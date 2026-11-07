"""Core data ingestion, Claude 3.5 Sonnet schema synthesis, and Excel/ZIP normalization engine.

Implements Rules A through F:
- Rule A: Leading-zero safe ingestion (`dtype=str`) for `.xlsx` and `.csv`.
- Rule B: Schema synthesis, drift prevention, full superset retention, and collision safeguards.
- Rule C: Cardinality-aware sampling (<= 150 unique sent in full; > 150 top 10 + annotation).
- Rule D: Whitespace, non-breaking space (`\\xa0`), and case normalization (`normalize_key`).
- Rule E: Numeric spec cleaning, commercial range preservation, and native Python `int`/`float` typing.
- Rule F: Visual exception flagging (`*` suffix + `#FFFF00` solid yellow fill) and column auto-fitting.
"""

from __future__ import annotations

import io
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from anthropic import Anthropic, NotFoundError
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from rules_manager import merge_synthesis_into_rules, normalize_key

# Solid yellow highlight required by Rule F (#FFFF00)
YELLOW_EXCEPTION_FILL = PatternFill(
    start_color="FFFF00",
    end_color="FFFF00",
    fill_type="solid",
)

HEADER_FILL = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid",
)
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
DATA_FONT = Font(name="Calibri", size=11)

# Cardinality threshold required by Rule C
CARDINALITY_THRESHOLD = 150
HIGH_CARDINALITY_SAMPLE_SIZE = 10

# Keywords identifying free-text, identifier, or numeric columns that must NEVER
# undergo categorical value_transforms (Rule C safeguard)
NON_CATEGORICAL_KEYWORDS = {
    "sku",
    "upc",
    "ean",
    "gtin",
    "barcode",
    "part_number",
    "part_no",
    "item_number",
    "item_no",
    "model_number",
    "model_no",
    "catalog_number",
    "mpn",
    "asin",
    "id",
    "title",
    "name",
    "description",
    "desc",
    "notes",
    "comment",
    "url",
    "link",
    "image",
    "bin",
    "warehouse_bin",
    "location_bin",
    "aisle",
    "shelf",
    "price",
    "cost",
    "msrp",
    "map",
    "wholesale",
    "retail",
    "watt",
    "wattage",
    "lumen",
    "lumens",
    "cri",
    "lifespan",
    "hours",
    "qty",
    "quantity",
    "stock",
    "inventory",
    "weight",
    "length",
    "width",
    "height",
    "diameter",
    "depth",
    "dimension",
    "beam_angle",
    "case_pack",
    "moq",
}


# ---------------------------------------------------------------------------
# Rule A: Ingestion & Leading-Zero Safety
# ---------------------------------------------------------------------------
def read_vendor_file(file_obj: Any, filename: str | None = None) -> pd.DataFrame:
    """Ingest a vendor `.xlsx` or `.csv` file with `dtype=str` to preserve leading zeros (Rule A).

    Barcodes, UPCs, and SKUs (e.g., "00492", "08500123") are kept strictly as strings.
    """
    if filename:
        name = filename
    elif isinstance(file_obj, (str, Path)):
        name = str(file_obj)
    else:
        name = getattr(file_obj, "name", "vendor_file.csv")
    ext = Path(name).suffix.lower()

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    if ext in {".xlsx", ".xls"}:
        df = pd.read_excel(file_obj, dtype=str)
    elif ext == ".csv":
        try:
            df = pd.read_csv(file_obj, dtype=str, keep_default_na=False)
        except UnicodeDecodeError:
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)
            df = pd.read_csv(
                file_obj,
                dtype=str,
                 encoding="latin-1",
                keep_default_na=False,
            )
    else:
        raise ValueError(f"Unsupported file format '{ext}' for '{name}'. Please upload .xlsx or .csv files.")

    # Replace any NaN/None values with empty strings while keeping all values as str
    df = df.fillna("")
    for col in df.columns:
        df[col] = df[col].astype(str).map(lambda x: "" if x.lower() == "nan" else x)

    # Clean column header names of leading/trailing whitespace
    df.columns = [str(c).strip() for c in df.columns]

    # Drop completely empty rows
    non_empty_mask = df.apply(lambda row: any(str(v).strip() != "" for v in row), axis=1)
    df = df.loc[non_empty_mask].reset_index(drop=True)

    return df


# ---------------------------------------------------------------------------
# Rule C: Cardinality-Aware Sampling (No 10-Row Blindspots)
# ---------------------------------------------------------------------------
def build_column_profiles(
    vendor_dfs: dict[str, pd.DataFrame],
    cardinality_threshold: int = CARDINALITY_THRESHOLD,
    sample_size: int = HIGH_CARDINALITY_SAMPLE_SIZE,
) -> dict[str, dict[str, Any]]:
    """Inspect every column across all uploaded files prior to calling Claude (Rule C).

    - If a column has <= 150 unique non-empty strings, captures and sends ALL unique values.
    - If a column has > 150 unique values, sends only the top 10 samples and annotates
      with `"(high cardinality, total X distinct)"`.
    """
    profiles: dict[str, dict[str, Any]] = {}

    for filename, df in vendor_dfs.items():
        file_profile: dict[str, Any] = {}
        for col in df.columns:
            # Extract non-empty values preserving order of appearance
            seen: set[str] = set()
            unique_vals: list[str] = []
            for raw_val in df[col]:
                cleaned_val = str(raw_val).replace("\xa0", " ").strip()
                if cleaned_val and cleaned_val not in seen:
                    seen.add(cleaned_val)
                    unique_vals.append(cleaned_val)

            distinct_count = len(unique_vals)
            if distinct_count <= cardinality_threshold:
                file_profile[col] = {
                    "distinct_count": distinct_count,
                    "cardinality": "low",
                    "values": unique_vals,
                }
            else:
                file_profile[col] = {
                    "distinct_count": distinct_count,
                    "cardinality": f"(high cardinality, total {distinct_count} distinct)",
                    "values": unique_vals[:sample_size],
                }
        profiles[filename] = file_profile

    return profiles


# ---------------------------------------------------------------------------
# Rule E: Numeric Spec Cleaning & Range Support
# ---------------------------------------------------------------------------
# Pre-compiled patterns for numeric specs and commercial ranges
_RANGE_HYPHEN_RE = re.compile(
    r"^\s*[$€£¥]?\s*(\d+(?:\.\d+)?)\s*(?:[a-zA-Z%°\"']+\s*)?[-–—]\s*[$€£¥]?\s*(\d+(?:\.\d+)?)\s*(?:[a-zA-Z%°/.\s\"']+)?$"
)
_RANGE_TO_RE = re.compile(
    r"^\s*[$€£¥]?\s*(\d+(?:\.\d+)?)\s*(?:[a-zA-Z%°\"']+\s*)?\bto\b\s*[$€£¥]?\s*(\d+(?:\.\d+)?)\s*(?:[a-zA-Z%°/.\s\"']+)?$",
    re.IGNORECASE,
)
_MULTI_SLASH_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*[a-zA-Z%°\"']*(?:\s*/\s*\d+(?:\.\d+)?\s*[a-zA-Z%°\"']*)+\s*[a-zA-Z%°\s\"']*$"
)
_SINGLE_NUMERIC_RE = re.compile(
    r"^\s*[$€£¥]?\s*([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?|[+-]?\.\d+)\s*\+?\s*(?:[a-zA-Z%°Ωµμ/.\s\"']+)?$"
)


def _to_python_number(num_str: str) -> int | float:
    """Convert a clean numeric string to native Python `int` or `float`."""
    cleaned = num_str.replace(",", "").strip()
    val = float(cleaned)
    if val.is_integer():
        return int(val)
    return val


def clean_numeric_spec(val: Any) -> int | float | str:
    """Clean a numeric specification or price value according to Rule E.

    - Single values: Strip unit suffixes & currency symbols (e.g., "60 W" -> 60, "$14.99" -> 14.99)
      and return native Python `int` or `float`.
    - Commercial spec ranges: Preserve valid ranges without units
      (e.g., "120-277V" -> "120-277", "30/70/100W" -> "30/70/100").
    - Empty values: Return `""`.
    - Unparseable text: Keep raw text and append `*` (e.g., "See Notes" -> "See Notes*").
    """
    if val is None:
        return ""
    if isinstance(val, bool):
        return f"{val}*"
    if isinstance(val, int):
        return val
    if isinstance(val, float):
        if pd.isna(val):
            return ""
        return int(val) if val.is_integer() else val

    raw_str = str(val).replace("\xa0", " ").strip()
    if not raw_str:
        return ""

    # 1. Check commercial hyphenated ranges (e.g., "120-277V", "110 - 240 VAC")
    m_hyphen = _RANGE_HYPHEN_RE.match(raw_str)
    if m_hyphen:
        n1, n2 = m_hyphen.group(1), m_hyphen.group(2)
        return f"{n1}-{n2}"

    m_to = _RANGE_TO_RE.match(raw_str)
    if m_to:
        n1, n2 = m_to.group(1), m_to.group(2)
        return f"{n1}-{n2}"

    # 2. Check multi-tap / slash ranges (e.g., "30/70/100W", "120/277V", "2700K/3000K/4000K")
    if "/" in raw_str and _MULTI_SLASH_RE.match(raw_str):
        numbers = re.findall(r"\d+(?:\.\d+)?", raw_str)
        if len(numbers) >= 2:
            return "/".join(numbers)

    # 3. Check single numeric value with optional currency/units (e.g., "60 W", "$14.99", "1,200 lm")
    m_single = _SINGLE_NUMERIC_RE.match(raw_str)
    if m_single:
        num_token = m_single.group(1)
        try:
            return _to_python_number(num_token)
        except ValueError:
            pass

    # 4. Unparseable text in a numeric column -> preserve raw text and append '*'
    if raw_str.endswith("*"):
        return raw_str
    return f"{raw_str}*"


# ---------------------------------------------------------------------------
# Rule B & C: Safeguards & Deduplication
# ---------------------------------------------------------------------------
def is_likely_non_categorical(column_name: str) -> bool:
    """Return True if `column_name` represents an identifier, free-text, or numeric field."""
    norm = normalize_key(column_name).replace(" ", "_")
    tokens = set(re.split(r"[_\s-]+", norm))
    if norm in NON_CATEGORICAL_KEYWORDS:
        return True
    if tokens & NON_CATEGORICAL_KEYWORDS:
        # Allow exceptions like "voltage_type" or "color_temperature"
        if "type" in tokens or "rating" in tokens or "color" in tokens or "finish" in tokens or "shape" in tokens:
            return False
        return True
    return False


def enforce_schema_and_collision_safeguards(
    vendor_dfs: dict[str, pd.DataFrame],
    current_rules: dict[str, Any],
    raw_plan: dict[str, Any],
) -> dict[str, Any]:
    """Deterministically enforce Rule B and Rule C on the synthesis plan.

    1. Drift prevention: Existing `master_columns` from `current_rules` stay first in exact order.
    2. Full superset: Every column from every file is mapped; unmapped columns are preserved.
    3. Collision safeguard: Two source columns from the same file NEVER map to the same master column.
       If a collision occurs, disambiguates or appends `_2`, `_3`, etc.
    4. Categorical safeguard: High-cardinality, numeric, and identifier/free-text columns are excluded
       from `categorical_columns` and `value_transforms`.
    """
    existing_master: list[str] = list(current_rules.get("master_columns", []))
    master_columns: list[str] = list(existing_master)
    master_lower_map: dict[str, str] = {c.lower(): c for c in master_columns}

    def register_master_col(col_name: str) -> str:
        clean_name = str(col_name).strip().replace(" ", "_")
        if not clean_name:
            clean_name = "Unnamed_Column"
        if clean_name.lower() in master_lower_map:
            return master_lower_map[clean_name.lower()]
        master_columns.append(clean_name)
        master_lower_map[clean_name.lower()] = clean_name
        return clean_name

    # Register Claude's proposed master columns in order (after existing master columns)
    for proposed_col in raw_plan.get("master_columns", []):
        if str(proposed_col).strip():
            register_master_col(str(proposed_col).strip())

    raw_file_mappings = raw_plan.get("file_mappings", {})
    if not isinstance(raw_file_mappings, dict):
        raw_file_mappings = {}

    known_header_map = {
        normalize_key(k): v for k, v in current_rules.get("header_mappings", {}).items()
    }
    for k, v in raw_plan.get("header_mappings", {}).items():
        if str(k).strip() and str(v).strip():
            known_header_map[normalize_key(k)] = str(v).strip()

    clean_file_mappings: dict[str, dict[str, str]] = {}

    for filename, df in vendor_dfs.items():
        proposed_for_file = raw_file_mappings.get(filename, {})
        if not isinstance(proposed_for_file, dict):
            proposed_for_file = {}
        # Build normalized lookup for proposed file mappings
        proposed_norm = {normalize_key(k): str(v).strip() for k, v in proposed_for_file.items() if str(v).strip()}

        used_targets_in_file: set[str] = set()
        file_map: dict[str, str] = {}

        for src_col in df.columns:
            norm_src = normalize_key(src_col)
            target = (
                proposed_for_file.get(src_col)
                or proposed_norm.get(norm_src)
                or known_header_map.get(norm_src)
                or src_col.strip().replace(" ", "_")
            )
            canonical_target = register_master_col(target)

            # Rule B Column Collision Safeguard:
            # Never allow two source columns in the same file to map to the same master column
            if canonical_target.lower() in used_targets_in_file:
                # Attempt semantic disambiguation using source header name first, else suffix _2, _3...
                sanitized_src = re.sub(r"[^a-zA-Z0-9]+", "_", src_col.strip()).strip("_")
                candidate = sanitized_src if sanitized_src and sanitized_src.lower() != canonical_target.lower() else f"{canonical_target}_2"
                suffix_idx = 2
                while candidate.lower() in used_targets_in_file:
                    suffix_idx += 1
                    candidate = f"{canonical_target}_{suffix_idx}"
                canonical_target = register_master_col(candidate)

            used_targets_in_file.add(canonical_target.lower())
            file_map[src_col] = canonical_target

        clean_file_mappings[filename] = file_map

    # Numeric columns consolidation
    numeric_set_lower: set[str] = set()
    numeric_columns: list[str] = []
    for col in list(current_rules.get("numeric_columns", [])) + list(raw_plan.get("numeric_columns", [])):
        col_str = str(col).strip()
        if not col_str:
            continue
        canonical = master_lower_map.get(col_str.lower(), col_str)
        if canonical.lower() not in numeric_set_lower:
            numeric_set_lower.add(canonical.lower())
            numeric_columns.append(canonical)

    # Categorical columns consolidation (excluding numeric & non-categorical identifier/text columns)
    categorical_set_lower: set[str] = set()
    categorical_columns: list[str] = []
    for col in list(current_rules.get("categorical_columns", [])) + list(raw_plan.get("categorical_columns", [])):
        col_str = str(col).strip()
        if not col_str:
            continue
        canonical = master_lower_map.get(col_str.lower(), col_str)
        if canonical.lower() in numeric_set_lower or is_likely_non_categorical(canonical):
            continue
        if canonical.lower() not in categorical_set_lower:
            categorical_set_lower.add(canonical.lower())
            categorical_columns.append(canonical)

    # Value transforms consolidation
    value_transforms: dict[str, dict[str, str]] = {}
    for source_dict in (current_rules.get("value_transforms", {}), raw_plan.get("value_transforms", {})):
        if not isinstance(source_dict, dict):
            continue
        for col_name, mapping in source_dict.items():
            if not isinstance(mapping, dict):
                continue
            col_str = str(col_name).strip()
            canonical = master_lower_map.get(col_str.lower(), col_str)
            if canonical.lower() in numeric_set_lower or is_likely_non_categorical(canonical):
                continue
            if canonical.lower() not in categorical_set_lower:
                categorical_set_lower.add(canonical.lower())
                categorical_columns.append(canonical)
            value_transforms.setdefault(canonical, {})
            for raw_k, target_v in mapping.items():
                norm_k = normalize_key(raw_k)
                target_str = str(target_v).strip()
                if norm_k and target_str and not target_str.endswith("*"):
                    value_transforms[canonical][norm_k] = target_str

    # Build consolidated header_mappings across all files
    consolidated_header_mappings: dict[str, str] = dict(known_header_map)
    for _fname, fmap in clean_file_mappings.items():
        for src_col, target_col in fmap.items():
            consolidated_header_mappings[normalize_key(src_col)] = target_col

    return {
        "master_columns": master_columns,
        "file_mappings": clean_file_mappings,
        "header_mappings": consolidated_header_mappings,
        "numeric_columns": numeric_columns,
        "categorical_columns": categorical_columns,
        "value_transforms": value_transforms,
    }


# ---------------------------------------------------------------------------
# Claude Schema & Taxonomy Synthesis (Resilient JSON Parser)
# ---------------------------------------------------------------------------
def _sanitize_json_text(raw_text: str) -> str:
    """Clean markdown fences, unescaped control/quote characters, and trailing/missing commas."""
    text = raw_text.replace("\xa0", " ").strip()

    # 1. Strip markdown code fences (even if unclosed)
    text = re.sub(r"^```(?:json|javascript|js)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```\s*$", "", text)

    # Isolate from first '{'
    first_brace = text.find("{")
    if first_brace != -1:
        text = text[first_brace:]
    last_brace = text.rfind("}")
    if last_brace != -1 and last_brace > 0:
        text = text[: last_brace + 1]

    # 2. State-machine pass: clean unescaped control chars, invalid escapes, and unescaped inner quotes
    out_chars: list[str] = []
    in_string = False
    i = 0
    n = len(text)
    valid_escapes = {'"', "\\", "/", "b", "f", "n", "r", "t", "u"}

    while i < n:
        ch = text[i]
        if in_string:
            if ch == "\\":
                if i + 1 < n:
                    nxt = text[i + 1]
                    if nxt in valid_escapes:
                        out_chars.append(ch)
                        out_chars.append(nxt)
                        i += 2
                        continue
                    else:
                        # Escape stray backslash
                        out_chars.append("\\\\")
                        i += 1
                        continue
                else:
                    out_chars.append("\\\\")
                    i += 1
                    continue
            elif ch == '"':
                # Look ahead past whitespace to see if this quote actually closes the JSON string
                j = i + 1
                while j < n and text[j] in {" ", "\t", "\r", "\n"}:
                    j += 1
                next_sig = text[j] if j < n else ""
                # A closing string quote in JSON must be followed by ':', ',', '}', ']', or EOF,
                # OR if followed by '"' on a new line (missing comma case)
                newline_before_next = "\n" in text[i + 1 : j]
                if next_sig in {":", ",", "}", "]", ""} or (next_sig == '"' and newline_before_next):
                    in_string = False
                    out_chars.append('"')
                else:
                    # Unescaped inner quote (e.g. 4" spotlight) -> escape it
                    out_chars.append('\\"')
                i += 1
                continue
            elif ch == "\n":
                out_chars.append("\\n")
                i += 1
                continue
            elif ch == "\r":
                out_chars.append("\\r")
                i += 1
                continue
            elif ch == "\t":
                out_chars.append("\\t")
                i += 1
                continue
            elif ord(ch) < 0x20:
                out_chars.append(" ")
                i += 1
                continue
            else:
                out_chars.append(ch)
                i += 1
                continue
        else:
            # Outside string: strip single-line // comments
            if ch == "/" and i + 1 < n and text[i + 1] == "/":
                while i < n and text[i] != "\n":
                    i += 1
                continue
            # Strip /* ... */ comments
            if ch == "/" and i + 1 < n and text[i + 1] == "*":
                i += 2
                while i + 1 < n and not (text[i] == "*" and text[i + 1] == "/"):
                    i += 1
                i += 2
                continue
            if ch == '"':
                in_string = True
            out_chars.append(ch)
            i += 1

    cleaned = "".join(out_chars)

    # 3. Insert missing commas between adjacent values/entries on consecutive lines
    # e.g. `"val1"\n  "key2":` or `}\n  "key2":` or `]\n  "key2":`
    cleaned = re.sub(
        r'(["}\]\d]|true|false|null)[ \t]*\r?\n([ \t]*["{\[])',
        r"\1,\n\2",
        cleaned,
    )

    # 4. Remove trailing commas before closing '}' or ']'
    cleaned = re.sub(r",(\s*[}\]])", r"\1", cleaned)

    # 5. If truncated mid-stream, close open string and balance open brackets/braces
    if in_string:
        cleaned += '"'
    stack: list[str] = []
    in_s = False
    esc = False
    for c in cleaned:
        if in_s:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_s = False
        else:
            if c == '"':
                in_s = True
            elif c == "{":
                stack.append("}")
            elif c == "[":
                stack.append("]")
            elif c in {"}", "]"} and stack and stack[-1] == c:
                stack.pop()
    if stack:
        cleaned = re.sub(r",\s*$", "", cleaned)
        cleaned += "".join(reversed(stack))
        cleaned = re.sub(r",(\s*[}\]])", r"\1", cleaned)

    return cleaned


def _extract_json_payload(text: str) -> dict[str, Any]:
    """Resiliently extract and parse a JSON object from Claude's response text."""
    if not text or not text.strip():
        raise ValueError("Received an empty response from Claude.")

    # Fast path: direct parse if already clean
    stripped = text.strip()
    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        pass

    # Resilient sanitization path
    sanitized = _sanitize_json_text(stripped)
    try:
        return json.loads(sanitized)
    except json.JSONDecodeError as err:
        # Final fallback: remove any remaining non-printable ASCII characters and retry
        fallback = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", sanitized)
        fallback = re.sub(r",(\s*[}\]])", r"\1", fallback)
        try:
            return json.loads(fallback)
        except json.JSONDecodeError:
            raise ValueError(
                f"Failed to parse JSON from Claude response ({err}). "
                f"Context around char {err.pos}: {sanitized[max(0, err.pos - 80):min(len(sanitized), err.pos + 80)]!r}"
            ) from err


def synthesize_schema_with_claude(
    vendor_dfs: dict[str, pd.DataFrame],
    current_rules: dict[str, Any],
    api_key: str | None = None,
    model: str | None = None,
) -> dict[str, Any]:
    """Use Anthropic Claude to synthesize the Master Schema and value transforms.

    Respects:
    - Rule B: Existing `master_columns` priority, Full Superset retention, and Zero Column Collisions.
    - Rule C: Cardinality-Aware Sampling (<= 150 distinct sent in full, > 150 top 10 + annotation).
    """
    resolved_key = (api_key or os.getenv("ANTHROPIC_API_KEY", "")).strip()
    if not resolved_key or resolved_key == "your_anthropic_api_key_here":
        raise ValueError(
            "Anthropic API key is missing. Please set ANTHROPIC_API_KEY in .env or enter it in the sidebar."
        )

    column_profiles = build_column_profiles(vendor_dfs)
    existing_master = current_rules.get("master_columns", [])

    system_prompt = (
        "You are an expert retail data architect specializing in multi-vendor product catalog harmonization "
        "(electrical supplies, lighting, hardware, and commercial inventory).\n"
        "You must output strictly valid RFC 8259 JSON and nothing else. "
        "Do not wrap the JSON in markdown code fences, do not include trailing commas, and ensure all "
        "double quotes inside string values are properly escaped."
    )

    # Provide a concise summary of current rules so the prompt and response stay compact
    rules_context = {
        "master_columns": existing_master,
        "header_mappings": current_rules.get("header_mappings", {}),
        "numeric_columns": current_rules.get("numeric_columns", []),
        "categorical_columns": current_rules.get("categorical_columns", []),
        "existing_value_transform_columns": list(current_rules.get("value_transforms", {}).keys()),
    }

    user_prompt = f"""Analyze the following uploaded vendor spreadsheet profiles and current persistent rulebook to produce a unified harmonization plan.

### CORE RULES YOU MUST FOLLOW:
1. **Schema Drift Prevention (Rule B)**:
   - Existing `master_columns` in the rulebook: {json.dumps(existing_master)}
   - If `existing_master` is non-empty, you MUST keep every column in `existing_master` in its exact order as the base standard, and only append genuinely new product attributes to the end of `master_columns`.
   - If `existing_master` is empty (first run), dynamically synthesize a logical, clean `Pascal_Snake_Case` Master Schema (e.g., `SKU`, `UPC`, `Brand`, `Product_Title`, `Wattage`, `Lumens`, `Color_Temperature`, `Base_Type`, `Bulb_Shape`, `Finish`, `Dimmable`, `Voltage`, `CRI`, `Lifespan_Hours`, `Wholesale_Price`, `MSRP_Price`, `Stock_Qty`, etc.).

2. **Full Superset Retention (Rule B)**:
   - Retain EVERY unique attribute across all uploaded vendor files. Even if only one vendor has a specialty column (e.g., `Warehouse_Bin` or `Lead_Time_Days`), include it in `master_columns`.

3. **Zero Intra-File Column Collisions (Rule B)**:
   - Two source columns from the same file must NEVER map to the same master column!
   - For example, if a vendor file has both "MSRP" and "Wholesale Net", map them to distinct master columns like `"MSRP_Price"` and `"Wholesale_Price"`.

4. **Cardinality-Aware Value Transforms (Rule C)**:
   - Generate `value_transforms` ONLY for low-cardinality categorical classification columns (e.g., `Color_Temperature`, `Base_Type`, `Bulb_Shape`, `Finish`, `Dimmable`, `Technology`, `Location_Rating`).
   - NEVER generate `value_transforms` for high-cardinality columns, SKUs, UPCs, part numbers, product titles, long descriptions, warehouse bins, or numeric columns.
   - In `value_transforms`, map lowercase normalized raw values to clean canonical standards (e.g., in `Color_Temperature`: `"soft white": "2700K"`, `"warm white": "3000K"`, `"cool white": "4000K"`, `"daylight": "5000K"`; in `Base_Type`: `"medium": "E26"`, `"med": "E26"`, `"candelabra": "E12"`; in `Finish`: `"frost": "Frosted"`, `"milky": "Opal"`, `"clear": "Clear"`).
   - Only include `value_transforms` for values actually observed in the uploaded files (pre-existing rules in `rules.json` are preserved automatically).
   - If a value in a categorical column is genuinely unrecognizable, ambiguous, or custom/non-standard (e.g., `"Custom White"`, `"Custom Sunset Glow"`, `"Full Spectrum UV"`, `"Prismatic Dichroic"`, `"Hand-Painted Gold"`, `"Bayonet B22d"`, `"Submersible IP68"`, `"TBD"`), DO NOT include it in `value_transforms` so the downstream engine flags it with `*` and highlights it in yellow.

5. **Numeric Columns Classification (Rule E)**:
   - List all master columns representing numeric measurements, electrical specs, counts, dimensions, or prices in `numeric_columns` (e.g., `Wattage`, `Lumens`, `Voltage`, `CRI`, `Lifespan_Hours`, `Wholesale_Price`, `MSRP_Price`, `Case_Quantity`, `Stock_Qty`, etc.).
   - Do NOT put `SKU`, `UPC`, `Barcode`, or `Part_Number` in `numeric_columns` (they are string identifiers with leading zeros).

### CURRENT PERSISTENT RULEBOOK SUMMARY:
{json.dumps(rules_context, indent=2)}

### UPLOADED VENDOR COLUMN PROFILES (Cardinality-Aware):
{json.dumps(column_profiles, indent=2)}

### REQUIRED JSON RESPONSE FORMAT:
Return ONLY a single valid JSON object matching this exact structure (no markdown fences, no trailing commas):
{{
  "master_columns": ["SKU", "Product_Title", "..."],
  "file_mappings": {{
    "filename1.csv": {{
      "Source Column A": "Master_Column_A",
      "Source Column B": "Master_Column_B"
    }}
  }},
  "numeric_columns": ["Wattage", "Lumens", "Voltage", "Wholesale_Price", "MSRP_Price"],
  "categorical_columns": ["Color_Temperature", "Base_Type", "Bulb_Shape", "Finish", "Dimmable"],
  "value_transforms": {{
    "Color_Temperature": {{
      "soft white": "2700K",
      "warm white": "3000K"
    }},
    "Base_Type": {{
      "medium": "E26",
      "candelabra": "E12"
    }}
  }}
}}"""

    client = Anthropic(api_key=resolved_key)
    preferred_model = model or os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5-5")
    candidate_models = [
        "claude-sonnet-5-5",
        preferred_model,
        "claude-3-5-sonnet-latest",
        "claude-3-5-sonnet-20241022",
        "claude-sonnet-4-20250514",
        "claude-3-5-haiku-latest",
    ]
    # Deduplicate candidate models preserving order
    unique_models: list[str] = []
    for m in candidate_models:
        if m and m not in unique_models:
            unique_models.append(m)

    last_error: Exception | None = None
    response_text = ""
    for candidate in unique_models:
        try:
            response = client.messages.create(
                model=candidate,
                max_tokens=8192,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
            )
            response_text = "".join(
                block.text for block in response.content if hasattr(block, "text")
            )
            break
        except NotFoundError as exc:
            last_error = exc
            continue

    if not response_text and last_error is not None:
        raise last_error

    raw_plan = _extract_json_payload(response_text)
    return enforce_schema_and_collision_safeguards(vendor_dfs, current_rules, raw_plan)


# ---------------------------------------------------------------------------
# Rule D, E, F: Deterministic Normalization & OpenPyXL Workbook Generation
# ---------------------------------------------------------------------------
def _resolve_categorical_transform_dict(
    master_col: str,
    value_transforms: dict[str, dict[str, str]],
) -> dict[str, str]:
    """Return normalized lookup dictionary for `master_col` (matching case-insensitively or by alias)."""
    col_lower = master_col.lower()
    raw_map: dict[str, str] = {}

    # Alias matching so e.g. "CCT" or "Color_Temp" inherits "Color_Temperature" transforms
    alias_groups = [
        ({"color_temperature", "color_temp", "cct", "kelvin", "light_color"}, "Color_Temperature"),
        ({"base_type", "base", "bulb_base", "socket", "cap_base"}, "Base_Type"),
        ({"finish", "bulb_finish", "lens_finish", "coating"}, "Finish"),
        ({"dimmable", "dimmability", "dimming"}, "Dimmable"),
        ({"bulb_shape", "shape", "lamp_shape", "envelope"}, "Bulb_Shape"),
        ({"technology", "light_source", "lamp_type", "bulb_type"}, "Technology"),
        ({"location_rating", "wet_rating", "environment", "indoor_outdoor"}, "Location_Rating"),
    ]
    for aliases, canonical_key in alias_groups:
        if col_lower in aliases and canonical_key in value_transforms:
            for k, v in value_transforms[canonical_key].items():
                raw_map[normalize_key(k)] = v

    for k_col, mapping in value_transforms.items():
        if k_col.lower() == col_lower and isinstance(mapping, dict):
            for k, v in mapping.items():
                raw_map[normalize_key(k)] = v

    # Also allow canonical target values themselves to match when normalized
    canonical_targets = list(set(raw_map.values()))
    for target in canonical_targets:
        norm_target = normalize_key(target)
        if norm_target and norm_target not in raw_map:
            raw_map[norm_target] = target

    return raw_map


def normalize_vendor_dataframe(
    df: pd.DataFrame,
    file_mapping: dict[str, str],
    synthesis_plan: dict[str, Any],
) -> tuple[pd.DataFrame, set[tuple[int, str]]]:
    """Transform a raw vendor DataFrame into the Master Schema layout.

    Returns:
        - `normalized_df`: DataFrame with exact `master_columns` in order.
          Numeric cells contain actual Python `int`/`float` (or range strings / flagged strings).
        - `flagged_cells`: Set of `(row_idx, master_col)` coordinates requiring solid yellow fill (#FFFF00).
    """
    master_columns: list[str] = synthesis_plan["master_columns"]
    numeric_cols_lower = {c.lower() for c in synthesis_plan.get("numeric_columns", [])}
    categorical_cols_lower = {c.lower() for c in synthesis_plan.get("categorical_columns", [])}
    value_transforms: dict[str, dict[str, str]] = synthesis_plan.get("value_transforms", {})

    # Invert `file_mapping` (src_col -> master_col) after collision safeguards
    master_to_src: dict[str, str] = {}
    for src_col, target_col in file_mapping.items():
        if src_col in df.columns:
            master_to_src[target_col] = src_col

    num_rows = len(df)
    output_data: dict[str, list[Any]] = {col: [""] * num_rows for col in master_columns}
    flagged_cells: set[tuple[int, str]] = set()

    for master_col in master_columns:
        src_col = master_to_src.get(master_col)
        if not src_col:
            continue

        col_lower = master_col.lower()
        is_numeric = col_lower in numeric_cols_lower
        is_categorical = (
            not is_numeric
            and not is_likely_non_categorical(master_col)
            and (col_lower in categorical_cols_lower or master_col in value_transforms)
        )
        cat_lookup = (
            _resolve_categorical_transform_dict(master_col, value_transforms)
            if is_categorical
            else {}
        )

        raw_series = df[src_col].tolist()
        transformed_col: list[Any] = []

        for row_idx, raw_val in enumerate(raw_series):
            # Clean non-breaking spaces while keeping leading zeros on raw strings
            raw_str = "" if raw_val is None else str(raw_val).replace("\xa0", " ").strip()
            if not raw_str:
                transformed_col.append("")
                continue

            if is_numeric:
                cleaned_num = clean_numeric_spec(raw_str)
                transformed_col.append(cleaned_num)
                if isinstance(cleaned_num, str) and cleaned_num.endswith("*"):
                    flagged_cells.add((row_idx, master_col))
            elif is_categorical:
                norm_key = normalize_key(raw_str)
                if norm_key in cat_lookup:
                    transformed_col.append(cat_lookup[norm_key])
                else:
                    flagged_val = raw_str if raw_str.endswith("*") else f"{raw_str}*"
                    transformed_col.append(flagged_val)
                    flagged_cells.add((row_idx, master_col))
            else:
                # Identifier / SKU / UPC / Title / Free-text column: preserve exact string
                transformed_col.append(raw_str)

        output_data[master_col] = transformed_col

    normalized_df = pd.DataFrame(output_data, columns=master_columns)
    return normalized_df, flagged_cells


def write_standardized_excel_bytes(
    normalized_df: pd.DataFrame,
    flagged_cells: set[tuple[int, str]],
    sheet_title: str = "Harmonized Catalog",
) -> bytes:
    """Write a normalized DataFrame to an in-memory OpenPyXL `.xlsx` file (Rules E & F).

    - Writes Python `int` and `float` as native numeric Excel cells so formulas work and no
      "Number stored as text" green warning triangles appear.
    - Preserves leading-zero identifiers (e.g., "00492") with explicit text format `@`.
    - Fills every flagged cell (ending with `*` or in `flagged_cells`) with solid yellow `#FFFF00`.
    - Auto-fits all column widths so no text is truncated.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]

    master_columns = list(normalized_df.columns)

    # 1. Write Header Row (Row 1)
    for col_idx, col_name in enumerate(master_columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"

    # 2. Write Data Rows (Rows 2..N+1)
    for row_idx in range(len(normalized_df)):
        excel_row = row_idx + 2
        for col_idx, col_name in enumerate(master_columns, start=1):
            val = normalized_df.iat[row_idx, col_idx - 1]
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.font = DATA_FONT

            is_flagged = (row_idx, col_name) in flagged_cells or (
                isinstance(val, str) and val.endswith("*")
            )

            if isinstance(val, bool):
                cell.value = str(val)
            elif isinstance(val, int):
                cell.value = int(val)
                cell.number_format = "0"
            elif isinstance(val, float):
                cell.value = float(val)
                cell.number_format = "0.00" if "price" in col_name.lower() or "cost" in col_name.lower() else "0.##"
            else:
                str_val = "" if val is None else str(val)
                cell.value = str_val
                # Preserve leading-zero strings explicitly as text format '@'
                if str_val:
                    cell.data_type = "s"
                    cell.number_format = "@"

            if is_flagged:
                cell.fill = YELLOW_EXCEPTION_FILL

    # 3. Auto-fit all column widths (Rule F)
    for col_idx, col_name in enumerate(master_columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        for row_idx in range(len(normalized_df)):
            val = normalized_df.iat[row_idx, col_idx - 1]
            val_str = "" if val is None else str(val)
            if len(val_str) > max_len:
                max_len = len(val_str)
        # Add padding so bold headers and values are never truncated
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 65)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def process_batch_to_zip(
    vendor_dfs: dict[str, pd.DataFrame],
    synthesis_plan: dict[str, Any],
    current_rules: dict[str, Any] | None = None,
    persist_rules: bool = True,
) -> tuple[bytes, list[dict[str, Any]], dict[str, Any]]:
    """Run deterministic batch normalization across all N vendor files and package into an in-memory ZIP.

    Returns:
        - `zip_bytes`: Raw bytes of the `.zip` archive containing N standardized `.xlsx` files.
        - `file_summaries`: List of per-file metrics (filename, output_filename, rows, mapped_columns, flagged_cells).
        - `updated_rules`: The updated rulebook dictionary after merging the synthesis plan.
    """
    updated_rules = current_rules or {}
    if current_rules is not None:
        updated_rules = merge_synthesis_into_rules(
            current_rules=current_rules,
            synthesis_plan=synthesis_plan,
            persist=persist_rules,
        )

    zip_buffer = io.BytesIO()
    file_summaries: list[dict[str, Any]] = []
    used_zip_names: set[str] = set()

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for original_filename, df in vendor_dfs.items():
            file_map = synthesis_plan.get("file_mappings", {}).get(original_filename, {})
            norm_df, flagged_cells = normalize_vendor_dataframe(
                df=df,
                file_mapping=file_map,
                synthesis_plan=synthesis_plan,
            )
            xlsx_bytes = write_standardized_excel_bytes(
                normalized_df=norm_df,
                flagged_cells=flagged_cells,
            )

            stem = Path(original_filename).stem
            out_name = f"{stem}_harmonized.xlsx"
            counter = 2
            while out_name.lower() in used_zip_names:
                out_name = f"{stem}_harmonized_{counter}.xlsx"
                counter += 1
            used_zip_names.add(out_name.lower())

            zf.writestr(out_name, xlsx_bytes)

            file_summaries.append(
                {
                    "Source File": original_filename,
                    "Harmonized File": out_name,
                    "Rows Processed": len(norm_df),
                    "Source Columns Mapped": len(file_map),
                    "Master Schema Columns": len(norm_df.columns),
                    "Exceptions Flagged (*)": len(flagged_cells),
                }
            )

    zip_buffer.seek(0)
    return zip_buffer.getvalue(), file_summaries, updated_rules
