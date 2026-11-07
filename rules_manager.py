"""Persistent rulebook manager for the Product Catalog Harmonizer.

Handles loading, saving, validating, normalizing, and resetting `rules.json`,
as well as merging newly synthesized schemas and taxonomy mappings without
causing schema drift across runs.
"""

from __future__ import annotations

import copy
import json
import os
import re
from pathlib import Path
from typing import Any

RULES_FILE_PATH = Path(__file__).resolve().parent / "rules.json"


def normalize_key(text: Any) -> str:
    """Normalize a string key for consistent dictionary lookups (Rule D).

    - Replaces non-breaking web spaces (\\xa0, \\u2007, \\u202f) with regular spaces.
    - Collapses multiple whitespace characters into a single space.
    - Converts strings to lowercase and trims leading/trailing spaces.
    """
    if text is None:
        return ""
    raw = str(text)
    # Replace non-breaking spaces and zero-width spaces with regular space
    raw = raw.replace("\xa0", " ").replace("\u2007", " ").replace("\u202f", " ")
    # Collapse all whitespace sequences (spaces, tabs, newlines) into a single space
    raw = re.sub(r"\s+", " ", raw)
    return raw.strip().lower()


DEFAULT_RULES: dict[str, Any] = {
    "master_columns": [],
    "header_mappings": {},
    "numeric_columns": [
        "Wattage",
        "Wattage_Equivalent",
        "Lumens",
        "Voltage",
        "CRI",
        "Lifespan_Hours",
        "Wholesale_Price",
        "MSRP_Price",
        "MAP_Price",
        "Unit_Price",
        "Cost_Price",
        "Case_Quantity",
        "Min_Order_Qty",
        "Stock_Qty",
        "Length_Inches",
        "Diameter_Inches",
        "Weight_Lbs",
        "Beam_Angle_Degrees",
    ],
    "categorical_columns": [
        "Color_Temperature",
        "Base_Type",
        "Bulb_Shape",
        "Finish",
        "Dimmable",
        "Technology",
        "Voltage_Type",
        "Location_Rating",
        "Certification",
    ],
    "value_transforms": {
        "Color_Temperature": {
            "amber": "2000K",
            "flame": "2000K",
            "candlelight": "2200K",
            "extra warm white": "2200K",
            "very warm white": "2400K",
            "soft white": "2700K",
            "warm": "2700K",
            "2700": "2700K",
            "2700k": "2700K",
            "2700 k": "2700K",
            "2700 kelvin": "2700K",
            "soft white (2700k)": "2700K",
            "warm white": "3000K",
            "bright warm white": "3000K",
            "3000": "3000K",
            "3000k": "3000K",
            "3000 k": "3000K",
            "3000 kelvin": "3000K",
            "warm white (3000k)": "3000K",
            "neutral white": "3500K",
            "natural white": "3500K",
            "3500": "3500K",
            "3500k": "3500K",
            "3500 k": "3500K",
            "cool white": "4000K",
            "bright white": "4000K",
            "4000": "4000K",
            "4000k": "4000K",
            "4000 k": "4000K",
            "4000 kelvin": "4000K",
            "cool white (4000k)": "4000K",
            "daylight": "5000K",
            "natural daylight": "5000K",
            "clear daylight": "5000K",
            "5000": "5000K",
            "5000k": "5000K",
            "5000 k": "5000K",
            "5000 kelvin": "5000K",
            "daylight (5000k)": "5000K",
            "cool daylight": "6500K",
            "daylight deluxe": "6500K",
            "6500": "6500K",
            "6500k": "6500K",
            "6500 k": "6500K",
        },
        "Base_Type": {
            "e26": "E26",
            "med": "E26",
            "medium": "E26",
            "medium screw": "E26",
            "medium base": "E26",
            "e26 medium": "E26",
            "e26 (medium)": "E26",
            "edison": "E26",
            "standard": "E26",
            "standard medium": "E26",
            "e12": "E12",
            "cand": "E12",
            "candelabra": "E12",
            "candelabra screw": "E12",
            "candelabra base": "E12",
            "e12 candelabra": "E12",
            "e12 (candelabra)": "E12",
            "e17": "E17",
            "intermediate": "E17",
            "intermediate base": "E17",
            "e39": "E39",
            "mogul": "E39",
            "mogul base": "E39",
            "mogul screw": "E39",
            "ex39": "EX39",
            "gu10": "GU10",
            "gu10 bi-pin": "GU10",
            "twist and lock": "GU10",
            "gu24": "GU24",
            "gu5.3": "GU5.3",
            "bi-pin": "G13",
            "medium bi-pin": "G13",
            "g13": "G13",
            "g4": "G4",
            "g8": "G8",
            "g9": "G9",
        },
        "Finish": {
            "clear": "Clear",
            "clr": "Clear",
            "transparent": "Clear",
            "frost": "Frosted",
            "frosted": "Frosted",
            "inside frost": "Frosted",
            "satin": "Frosted",
            "diffuse": "Frosted",
            "milky": "Opal",
            "milk": "Opal",
            "white": "Opal",
            "opal": "Opal",
            "opaque": "Opal",
            "soft white": "Opal",
            "amber": "Amber",
            "vintage amber": "Amber",
            "gold tint": "Amber",
            "silver bowl": "Silver Bowl",
        },
        "Dimmable": {
            "yes": "Yes",
            "y": "Yes",
            "true": "Yes",
            "1": "Yes",
            "dimmable": "Yes",
            "triac dimmable": "Yes",
            "0-10v dimmable": "Yes",
            "phase dimmable": "Yes",
            "no": "No",
            "n": "No",
            "false": "No",
            "0": "No",
            "non-dimmable": "No",
            "non dimmable": "No",
            "not dimmable": "No",
        },
        "Bulb_Shape": {
            "a19": "A19",
            "a-19": "A19",
            "standard a19": "A19",
            "a21": "A21",
            "a-21": "A21",
            "br30": "BR30",
            "br-30": "BR30",
            "br40": "BR40",
            "br-40": "BR40",
            "par20": "PAR20",
            "par30": "PAR30",
            "par38": "PAR38",
            "mr16": "MR16",
            "mr-16": "MR16",
            "b10": "B10",
            "b11": "B11",
            "candelabra": "B11",
            "torpedo": "B11",
            "candle": "B11",
            "ca10": "CA10",
            "flame tip": "CA10",
            "g16.5": "G16.5",
            "g25": "G25",
            "globe": "G25",
            "st19": "ST19",
            "st64": "ST19",
            "edison": "ST19",
            "t8": "T8",
            "t-8": "T8",
            "linear t8": "T8",
            "t5": "T5",
        },
        "Technology": {
            "led": "LED",
            "light emitting diode": "LED",
            "filament led": "LED",
            "incandescent": "Incandescent",
            "inc": "Incandescent",
            "incand": "Incandescent",
            "halogen": "Halogen",
            "hal": "Halogen",
            "cfl": "CFL",
            "compact fluorescent": "CFL",
            "fluorescent": "Fluorescent",
            "hid": "HID",
            "metal halide": "HID",
        },
        "Location_Rating": {
            "dry": "Dry",
            "dry location": "Dry",
            "indoor dry": "Dry",
            "damp": "Damp",
            "damp location": "Damp",
            "indoor/covered outdoor": "Damp",
            "wet": "Wet",
            "wet location": "Wet",
            "outdoor wet": "Wet",
            "weatherproof": "Wet",
        },
    },
}


def validate_rules_dict(data: Any) -> tuple[bool, str]:
    """Validate that a parsed JSON object conforms to the rulebook structure."""
    if not isinstance(data, dict):
        return False, "Root JSON element must be an object (dictionary)."

    required_list_keys = ["master_columns", "numeric_columns", "categorical_columns"]
    for key in required_list_keys:
        if key in data and not isinstance(data[key], list):
            return False, f"Key '{key}' must be a JSON list of strings."
        if key in data:
            for item in data[key]:
                if not isinstance(item, str):
                    return False, f"All elements in '{key}' must be strings."

    if "header_mappings" in data:
        if not isinstance(data["header_mappings"], dict):
            return False, "Key 'header_mappings' must be a JSON object mapping source headers to master columns."
        for k, v in data["header_mappings"].items():
            if not isinstance(k, str) or not isinstance(v, str):
                return False, "All keys and values in 'header_mappings' must be strings."

    if "value_transforms" in data:
        if not isinstance(data["value_transforms"], dict):
            return False, "Key 'value_transforms' must be a JSON object mapping column names to value dictionaries."
        for col_name, mapping in data["value_transforms"].items():
            if not isinstance(col_name, str) or not isinstance(mapping, dict):
                return (
                    False,
                    f"Value transform entry for '{col_name}' must be a JSON object.",
                )
            for raw_k, target_v in mapping.items():
                if not isinstance(raw_k, str) or not isinstance(target_v, str):
                    return (
                        False,
                        f"All keys and target values in 'value_transforms.{col_name}' must be strings.",
                    )

    return True, "Valid rulebook JSON."


def sanitize_rules(rules: dict[str, Any]) -> dict[str, Any]:
    """Ensure all required keys exist and normalize lookup keys per Rule D."""
    cleaned: dict[str, Any] = {
        "master_columns": [],
        "header_mappings": {},
        "numeric_columns": [],
        "categorical_columns": [],
        "value_transforms": {},
    }

    # Preserve master_columns order without duplicates
    seen_master: set[str] = set()
    for col in rules.get("master_columns", []):
        col_str = str(col).strip()
        if col_str and col_str not in seen_master:
            seen_master.add(col_str)
            cleaned["master_columns"].append(col_str)

    # Normalize header_mappings keys
    for src_header, target_col in rules.get("header_mappings", {}).items():
        norm_src = normalize_key(src_header)
        target_str = str(target_col).strip()
        if norm_src and target_str:
            cleaned["header_mappings"][norm_src] = target_str

    # Preserve numeric_columns without duplicates
    seen_num: set[str] = set()
    for col in rules.get("numeric_columns", []):
        col_str = str(col).strip()
        if col_str and col_str not in seen_num:
            seen_num.add(col_str)
            cleaned["numeric_columns"].append(col_str)

    # Preserve categorical_columns without duplicates
    seen_cat: set[str] = set()
    for col in rules.get("categorical_columns", []):
        col_str = str(col).strip()
        if col_str and col_str not in seen_cat:
            seen_cat.add(col_str)
            cleaned["categorical_columns"].append(col_str)

    # Normalize value_transforms keys
    for col_name, mapping in rules.get("value_transforms", {}).items():
        col_str = str(col_name).strip()
        if not col_str or not isinstance(mapping, dict):
            continue
        cleaned["value_transforms"].setdefault(col_str, {})
        for raw_val, canonical_val in mapping.items():
            norm_val = normalize_key(raw_val)
            canon_str = str(canonical_val).strip()
            if norm_val and canon_str:
                cleaned["value_transforms"][col_str][norm_val] = canon_str

    return cleaned


def load_rules(filepath: Path | str = RULES_FILE_PATH) -> dict[str, Any]:
    """Load `rules.json` from disk, initializing it with defaults if missing."""
    path = Path(filepath)
    if not path.exists():
        default_copy = sanitize_rules(copy.deepcopy(DEFAULT_RULES))
        save_rules(default_copy, path)
        return default_copy

    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        is_valid, _ = validate_rules_dict(data)
        if not is_valid:
            default_copy = sanitize_rules(copy.deepcopy(DEFAULT_RULES))
            return default_copy
        return sanitize_rules(data)
    except (json.JSONDecodeError, OSError):
        default_copy = sanitize_rules(copy.deepcopy(DEFAULT_RULES))
        return default_copy


def save_rules(rules: dict[str, Any], filepath: Path | str = RULES_FILE_PATH) -> dict[str, Any]:
    """Sanitize and persist the rulebook dictionary to `rules.json`."""
    path = Path(filepath)
    cleaned = sanitize_rules(rules)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(".json.tmp")
    with temp_path.open("w", encoding="utf-8") as f:
        json.dump(cleaned, f, indent=2, ensure_ascii=False)
        f.write("\n")
    os.replace(temp_path, path)
    return cleaned


def reset_rules(filepath: Path | str = RULES_FILE_PATH) -> dict[str, Any]:
    """Reset `rules.json` to the factory default taxonomy rulebook."""
    default_copy = copy.deepcopy(DEFAULT_RULES)
    return save_rules(default_copy, filepath)


def merge_synthesis_into_rules(
    current_rules: dict[str, Any],
    synthesis_plan: dict[str, Any],
    persist: bool = True,
    filepath: Path | str = RULES_FILE_PATH,
) -> dict[str, Any]:
    """Merge a Claude synthesis plan into the persistent rulebook without schema drift.

    Rule B Enforcement:
    - Existing `master_columns` always retain their exact order.
    - Genuinely new master columns from `synthesis_plan` are appended at the end.
    - All learned header mappings, column classifications, and value transforms are merged.
    """
    merged = sanitize_rules(copy.deepcopy(current_rules))

    # 1. Master columns: keep existing in order, append new ones
    existing_master_lower = {c.lower(): c for c in merged["master_columns"]}
    proposed_master = synthesis_plan.get("master_columns", [])
    for col in proposed_master:
        col_str = str(col).strip()
        if not col_str:
            continue
        if col_str.lower() not in existing_master_lower:
            merged["master_columns"].append(col_str)
            existing_master_lower[col_str.lower()] = col_str

    # Also ensure any target column referenced in file_mappings is in master_columns
    file_mappings = synthesis_plan.get("file_mappings", {})
    if isinstance(file_mappings, dict):
        for _filename, col_map in file_mappings.items():
            if not isinstance(col_map, dict):
                continue
            for src_col, target_col in col_map.items():
                target_str = str(target_col).strip()
                if not target_str:
                    continue
                canonical_target = existing_master_lower.get(target_str.lower(), target_str)
                if canonical_target.lower() not in existing_master_lower:
                    merged["master_columns"].append(canonical_target)
                    existing_master_lower[canonical_target.lower()] = canonical_target
                norm_src = normalize_key(src_col)
                if norm_src:
                    merged["header_mappings"][norm_src] = canonical_target

    # Merge explicit header_mappings if present in synthesis_plan
    for src_col, target_col in synthesis_plan.get("header_mappings", {}).items():
        norm_src = normalize_key(src_col)
        target_str = str(target_col).strip()
        if norm_src and target_str:
            canonical_target = existing_master_lower.get(target_str.lower(), target_str)
            if canonical_target.lower() not in existing_master_lower:
                merged["master_columns"].append(canonical_target)
                existing_master_lower[canonical_target.lower()] = canonical_target
            merged["header_mappings"][norm_src] = canonical_target

    # 2. Numeric columns
    existing_num_lower = {c.lower() for c in merged["numeric_columns"]}
    for col in synthesis_plan.get("numeric_columns", []):
        col_str = str(col).strip()
        if not col_str:
            continue
        canonical_col = existing_master_lower.get(col_str.lower(), col_str)
        if canonical_col.lower() not in existing_num_lower:
            merged["numeric_columns"].append(canonical_col)
            existing_num_lower.add(canonical_col.lower())

    # 3. Categorical columns
    existing_cat_lower = {c.lower() for c in merged["categorical_columns"]}
    for col in synthesis_plan.get("categorical_columns", []):
        col_str = str(col).strip()
        if not col_str:
            continue
        canonical_col = existing_master_lower.get(col_str.lower(), col_str)
        if canonical_col.lower() not in existing_cat_lower:
            merged["categorical_columns"].append(canonical_col)
            existing_cat_lower.add(canonical_col.lower())

    # 4. Value transforms
    existing_vt_keys_lower = {k.lower(): k for k in merged["value_transforms"]}
    for col_name, mapping in synthesis_plan.get("value_transforms", {}).items():
        if not isinstance(mapping, dict):
            continue
        col_str = str(col_name).strip()
        if not col_str:
            continue
        # Match to existing master column casing or existing value_transforms key
        canonical_col = existing_master_lower.get(
            col_str.lower(),
            existing_vt_keys_lower.get(col_str.lower(), col_str),
        )
        merged["value_transforms"].setdefault(canonical_col, {})
        existing_vt_keys_lower[canonical_col.lower()] = canonical_col

        for raw_val, target_val in mapping.items():
            norm_raw = normalize_key(raw_val)
            target_str = str(target_val).strip()
            # Ignore invalid or asterisked target transforms from LLM
            if norm_raw and target_str and not target_str.endswith("*"):
                merged["value_transforms"][canonical_col][norm_raw] = target_str

    if persist:
        return save_rules(merged, filepath)
    return sanitize_rules(merged)

