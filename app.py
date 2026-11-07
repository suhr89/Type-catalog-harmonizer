"""Streamlit UI for the Product Catalog Harmonizer.

Implements Rule G (Explicit 4-Stage State Machine: UPLOAD -> REVIEW -> GENERATE -> COMPLETE)
and Tab 2 (Persistent Rulebook Memory editor with JSON validation and reset capability).
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from normalizer import (
    process_batch_to_zip,
    read_vendor_file,
    synthesize_schema_with_claude,
)
from rules_manager import (
    RULES_FILE_PATH,
    load_rules,
    reset_rules,
    save_rules,
    validate_rules_dict,
)

# Load environment variables from .env
load_dotenv()

st.set_page_config(
    page_title="Product Catalog Harmonizer",
    page_icon="📦",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Session State Initialization (Rule G State Machine)
# ---------------------------------------------------------------------------
def init_session_state() -> None:
    """Initialize all session state variables required for the 4-stage state machine."""
    if "stage" not in st.session_state:
        st.session_state.stage = "UPLOAD"  # UPLOAD | REVIEW | GENERATE | COMPLETE
    if "vendor_dfs" not in st.session_state:
        st.session_state.vendor_dfs = {}
    if "synthesis_plan" not in st.session_state:
        st.session_state.synthesis_plan = None
    if "zip_bytes" not in st.session_state:
        st.session_state.zip_bytes = None
    if "file_summaries" not in st.session_state:
        st.session_state.file_summaries = []
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0
    if "rulebook_editor_version" not in st.session_state:
        st.session_state.rulebook_editor_version = 0


def reset_workflow_to_upload() -> None:
    """Reset the batch harmonization state machine back to the UPLOAD stage."""
    st.session_state.stage = "UPLOAD"
    st.session_state.vendor_dfs = {}
    st.session_state.synthesis_plan = None
    st.session_state.zip_bytes = None
    st.session_state.file_summaries = []
    st.session_state.uploader_key += 1


init_session_state()
current_rules = load_rules()


# ---------------------------------------------------------------------------
# Sidebar: Configuration & Persistent Schema Status
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Configuration")
    env_key = os.getenv("ANTHROPIC_API_KEY", "")
    default_key = "" if env_key == "your_anthropic_api_key_here" else env_key

    api_key_input = st.text_input(
        "Anthropic API Key",
        value=default_key,
        type="password",
        help="Loaded from .env if configured, or enter your Anthropic API key here.",
    )
    active_api_key = api_key_input.strip() or default_key

    model_name = st.selectbox(
        "Claude Model",
        options=[
            "claude-sonnet-5-5",
            "claude-3-5-sonnet-latest",
            "claude-3-5-sonnet-20241022",
            "claude-sonnet-4-20250514",
        ],
        index=0,
        help="Claude synthesizes the Master Schema and maps low-cardinality categorical values.",
    )

    st.divider()
    st.subheader("📚 Rulebook Snapshot")
    col_a, col_b = st.columns(2)
    col_a.metric("Master Columns", len(current_rules.get("master_columns", [])))
    col_b.metric("Header Rules", len(current_rules.get("header_mappings", {})))

    total_transforms = sum(
        len(v) for v in current_rules.get("value_transforms", {}).values() if isinstance(v, dict)
    )
    st.metric("Taxonomy Value Mappings", total_transforms)

    st.caption(
        f"Current Stage: **`{st.session_state.stage}`**\n\n"
        "Leading zeros on SKUs/UPCs are protected via `dtype=str`. "
        "Unmapped values are flagged with `*` and highlighted in yellow (`#FFFF00`)."
    )


# ---------------------------------------------------------------------------
# Main Application Header & Tabs
# ---------------------------------------------------------------------------
st.title("📦 Product Catalog Harmonizer")
st.markdown(
    "Standardize multi-vendor product spreadsheets (`.xlsx` and `.csv`) into a unified, drift-free "
    "**Master Schema** with automated unit cleaning, taxonomy normalization, and visual exception highlighting."
)

tab_harmonizer, tab_rulebook = st.tabs(
    ["🔄 Catalog Harmonizer", "📖 Rulebook Memory (`rules.json`)"]
)


# ===========================================================================
# TAB 1: CATALOG HARMONIZER (4-STAGE STATE MACHINE)
# ===========================================================================
with tab_harmonizer:
    stage = st.session_state.stage

    # -----------------------------------------------------------------------
    # STAGE 1: UPLOAD
    # -----------------------------------------------------------------------
    if stage == "UPLOAD":
        with st.expander("👋 New here? How this app works & who it is for", expanded=True):
            st.markdown(
                """
                **Who this app is for**
                Built for anyone managing fragmented vendor, supplier, or partner spreadsheets—not just lighting or electrical wholesalers, but **e-commerce resellers, wholesale distributors, marketplace aggregators, retail buyers, and inventory managers** across any industry.

                **How it reconciles & standardizes your data**
                * **Unified Master Schema:** Dynamically deduces a single, consistent column layout across all uploaded files—retaining every unique attribute while remembering your schema in `rules.json` so future runs never drift.
                * **Smart Unit & Spec Cleaning:** Strips unit suffixes and currency symbols (watts, prices, dimensions, counts) into native Excel numbers while preserving valid commercial spec ranges (e.g., `120-277`, `30/70/100`).
                * **Uniform Taxonomy Mapping:** Translates conflicting vendor terminology (e.g., *"Soft White"*, *"2700 Kelvin"*, *"Med Base"*) into clean, standardized categories.
                * **Leading-Zero Protection:** Reads all SKUs, UPCs, and barcodes strictly as text (`dtype=str`) so codes like `00492` or `085001234501` never lose their leading zeros.
                * **Visual Exception Flagging:** Any ambiguous categorical value or unparseable numeric note (e.g., *"Call for Price"*) is preserved, tagged with an asterisk (`*`), and highlighted in solid yellow (`#FFFF00`) so you can press `Ctrl + F` for `*` in Excel to review exceptions in seconds.

                **👇 Get started below**
                Drag and drop two or more vendor `.xlsx` or `.csv` spreadsheets into the **file uploader box directly below**, then click **🚀 Harmonize Vendor Catalogs**.
                """
            )

        st.subheader("1. Upload Vendor Spreadsheets")
        st.markdown(
            "Upload any number ($N$) of vendor catalog files (`.xlsx` or `.csv`). "
            "All files are ingested with `dtype=str` so UPCs, barcodes, and SKUs keep their leading zeros."
        )

        uploaded_files = st.file_uploader(
            "Drop vendor product sheets here (.xlsx, .csv)",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
            key=f"vendor_uploader_{st.session_state.uploader_key}",
        )

        audit_before_generate = st.checkbox(
            "Review mappings before generating final files",
            value=True,
            help="When checked, transitions to the REVIEW stage so you can inspect the Master Schema, header mappings, and value transforms before generating the ZIP.",
        )

        if st.button("🚀 Harmonize Vendor Catalogs", type="primary", disabled=not uploaded_files):
            if not active_api_key:
                st.error(
                    "Anthropic API Key is required. Please configure `ANTHROPIC_API_KEY` in `.env` or enter it in the sidebar."
                )
            else:
                try:
                    with st.spinner("Ingesting vendor spreadsheets with leading-zero protection (`dtype=str`)..."):
                        ingested_dfs: dict[str, pd.DataFrame] = {}
                        used_names: set[str] = set()
                        for uf in uploaded_files:
                            fname = uf.name
                            if fname in used_names:
                                idx = 2
                                stem, ext = os.path.splitext(fname)
                                while f"{stem}_{idx}{ext}" in used_names:
                                    idx += 1
                                fname = f"{stem}_{idx}{ext}"
                            used_names.add(fname)
                            ingested_dfs[fname] = read_vendor_file(uf, filename=fname)

                    with st.spinner(
                        "Synthesizing Master Schema and cardinality-aware value mappings with Claude 3.5 Sonnet..."
                    ):
                        latest_rules = load_rules()
                        plan = synthesize_schema_with_claude(
                            vendor_dfs=ingested_dfs,
                            current_rules=latest_rules,
                            api_key=active_api_key,
                            model=model_name,
                        )

                    st.session_state.vendor_dfs = ingested_dfs
                    st.session_state.synthesis_plan = plan

                    if audit_before_generate:
                        st.session_state.stage = "REVIEW"
                    else:
                        st.session_state.stage = "GENERATE"
                    st.rerun()

                except Exception as exc:
                    st.error(f"Error during catalog analysis: {exc}")

    # -----------------------------------------------------------------------
    # STAGE 2: REVIEW (Optional Audit Stage)
    # -----------------------------------------------------------------------
    elif stage == "REVIEW":
        st.subheader("2. Review Proposed Master Schema & Mappings")
        st.info(
            "Inspect the synthesized Master Schema, per-file column mappings, and categorical value transforms below. "
            "Click **Confirm & Generate** to build the standardized Excel batch, or **Cancel** to start over."
        )

        plan: dict[str, Any] = st.session_state.synthesis_plan or {}
        vendor_dfs: dict[str, pd.DataFrame] = st.session_state.vendor_dfs or {}

        action_col1, action_col2, _spacer = st.columns([1.5, 1, 4])
        with action_col1:
            if st.button("✅ Confirm & Generate", type="primary", use_container_width=True):
                st.session_state.stage = "GENERATE"
                st.rerun()
        with action_col2:
            if st.button("✖️ Cancel", use_container_width=True):
                reset_workflow_to_upload()
                st.rerun()

        st.divider()

        # 1. Master Schema Columns Overview
        master_cols = plan.get("master_columns", [])
        numeric_cols_lower = {c.lower() for c in plan.get("numeric_columns", [])}
        categorical_cols_lower = {c.lower() for c in plan.get("categorical_columns", [])}

        st.markdown(f"#### Unified Master Schema ({len(master_cols)} Columns)")
        schema_rows = []
        existing_saved_lower = {c.lower() for c in current_rules.get("master_columns", [])}
        for idx, col_name in enumerate(master_cols, start=1):
            c_low = col_name.lower()
            if c_low in numeric_cols_lower:
                col_type = "Numeric Spec / Price (Rule E)"
            elif c_low in categorical_cols_lower:
                col_type = "Categorical Taxonomy (Rule C/D)"
            else:
                col_type = "Identifier / Text (Rule A)"
            origin = "Existing Rulebook" if c_low in existing_saved_lower else "Newly Synthesized"
            schema_rows.append(
                {
                    "Column Order": idx,
                    "Master Column": col_name,
                    "Data Handling Type": col_type,
                    "Schema Origin": origin,
                }
            )
        st.dataframe(pd.DataFrame(schema_rows), use_container_width=True, hide_index=True)

        # 2. Per-File Source -> Master Header Mappings
        st.markdown("#### Source-to-Master Header Mappings by Vendor File")
        mapping_rows = []
        for fname, fmap in plan.get("file_mappings", {}).items():
            for src_col, target_col in fmap.items():
                mapping_rows.append(
                    {
                        "Vendor File": fname,
                        "Source Column Header": src_col,
                        "Mapped Master Column": target_col,
                    }
                )
        if mapping_rows:
            st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)

        # 3. Categorical Value Transforms
        st.markdown("#### Categorical Value Transforms (`value_transforms`)")
        transform_rows = []
        for col_name, val_map in plan.get("value_transforms", {}).items():
            if isinstance(val_map, dict):
                for raw_k, canon_v in val_map.items():
                    transform_rows.append(
                        {
                            "Master Column": col_name,
                            "Normalized Source Term": raw_k,
                            "Standardized Target Value": canon_v,
                        }
                    )
        if transform_rows:
            st.dataframe(pd.DataFrame(transform_rows), use_container_width=True, hide_index=True)
        else:
            st.caption("No categorical value transforms required for these columns.")

    # -----------------------------------------------------------------------
    # STAGE 3: GENERATE (Deterministic Batch Processing)
    # -----------------------------------------------------------------------
    elif stage == "GENERATE":
        st.subheader("3. Generating Standardized Excel Files...")
        try:
            with st.spinner("Normalizing rows, formatting Excel types, applying yellow flags, and building ZIP..."):
                latest_rules = load_rules()
                zip_bytes, summaries, _updated_rules = process_batch_to_zip(
                    vendor_dfs=st.session_state.vendor_dfs,
                    synthesis_plan=st.session_state.synthesis_plan,
                    current_rules=latest_rules,
                    persist_rules=True,
                )
                st.session_state.zip_bytes = zip_bytes
                st.session_state.file_summaries = summaries
                st.session_state.rulebook_editor_version += 1
                st.session_state.stage = "COMPLETE"
                st.rerun()
        except Exception as exc:
            st.error(f"Failed during Excel batch generation: {exc}")
            if st.button("⬅️ Back to Upload"):
                reset_workflow_to_upload()
                st.rerun()

    # -----------------------------------------------------------------------
    # STAGE 4: COMPLETE
    # -----------------------------------------------------------------------
    elif stage == "COMPLETE":
        st.success(
            "🎉 Batch Harmonization Complete! All vendor catalogs have been standardized to the unified Master Schema "
            "and your persistent `rules.json` taxonomy has been updated."
        )

        st.info(
            "💡 **Pro Tip for Exception Review in Excel:** Open any harmonized `.xlsx` file, press **`Ctrl + F`** "
            "(or `Cmd + F` on Mac), and search for **`*`**. Every low-confidence categorical term or unparseable numeric "
            "value is suffixed with `*` and highlighted in solid yellow (`#FFFF00`) so you can jump directly between exceptions."
        )

        summaries = st.session_state.file_summaries or []
        if summaries:
            total_files = len(summaries)
            total_rows = sum(item["Rows Processed"] for item in summaries)
            total_exceptions = sum(item["Exceptions Flagged (*)"] for item in summaries)
            master_col_count = summaries[0]["Master Schema Columns"]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Standardized Files", total_files)
            m2.metric("Total Products Processed", total_rows)
            m3.metric("Unified Master Columns", master_col_count)
            m4.metric("Flagged Exceptions (*)", total_exceptions)

            st.dataframe(pd.DataFrame(summaries), use_container_width=True, hide_index=True)

        st.divider()
        btn_col1, btn_col2, _spacer = st.columns([2, 1.5, 3])
        with btn_col1:
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 Download Harmonized Catalogs (.ZIP)",
                data=st.session_state.zip_bytes or b"",
                file_name=f"harmonized_catalogs_{timestamp_str}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
        with btn_col2:
            if st.button("🔄 Process Another Batch", use_container_width=True):
                reset_workflow_to_upload()
                st.rerun()


# ===========================================================================
# TAB 2: RULEBOOK MEMORY (VIEW, EDIT, VALIDATE, OR RESET `rules.json`)
# ===========================================================================
with tab_rulebook:
    st.subheader("📖 Persistent Taxonomy Rulebook (`rules.json`)")
    st.markdown(
        "This rulebook acts as the persistent memory across harmonization runs to prevent schema drift. "
        "You can directly inspect, edit, and save custom mappings below, or reset to factory defaults."
    )

    disk_rules = load_rules()
    formatted_rules_json = json.dumps(disk_rules, indent=2, ensure_ascii=False)

    edited_json_text = st.text_area(
        "Edit `rules.json` directly (JSON format):",
        value=formatted_rules_json,
        height=520,
        key=f"rules_json_editor_{st.session_state.rulebook_editor_version}",
    )

    r_col1, r_col2, _r_spacer = st.columns([1.5, 1.5, 4])
    with r_col1:
        if st.button("💾 Save Changes", type="primary", use_container_width=True):
            try:
                parsed_rules = json.loads(edited_json_text)
                is_valid, validation_msg = validate_rules_dict(parsed_rules)
                if not is_valid:
                    st.error(f"Schema validation failed: {validation_msg}")
                else:
                    save_rules(parsed_rules)
                    st.session_state.rulebook_editor_version += 1
                    st.success(f"Successfully validated and saved changes to `{RULES_FILE_PATH.name}`.")
                    st.rerun()
            except json.JSONDecodeError as json_err:
                st.error(f"Invalid JSON syntax: {json_err}")

    with r_col2:
        if st.button("🧹 Reset to Default", use_container_width=True):
            reset_rules()
            st.session_state.rulebook_editor_version += 1
            st.success("Rulebook has been reset to factory defaults (`master_columns` cleared).")
            st.rerun()

