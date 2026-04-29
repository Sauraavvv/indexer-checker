"""
app.py — Streamlit UI for URL Index Checker
Run with:  streamlit run app.py
"""

import asyncio
import os

import pandas as pd
import streamlit as st

from processing import run_sheets, OUTPUT_FILE, DEFAULT_MASTER_TAB

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="URL Index Checker",
    page_icon="🔍",
    layout="wide",
)

st.title("🔍 URL Index Checker")
st.caption("Checks Google Search indexing status for up to 5 sheets simultaneously.")

st.divider()


# ─────────────────────────────────────────────────────────────
# INPUT FORM
# ─────────────────────────────────────────────────────────────
with st.form("config_form"):
    st.subheader("📄 Sheet URLs  (leave blank to skip)")

    sheet_urls = []
    for i in range(5):
        url = st.text_input(
            f"Sheet {i + 1}",
            key=f"url_{i}",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
        sheet_urls.append(url)

    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        workers = st.number_input(
            "⚡ Number of Workers (parallel browser tabs)",
            min_value=1, max_value=20, value=4, step=1,
        )
    with col_b:
        master_tab = st.number_input(
            "📑 Master Tab Number (1-based, which tab has the URL list)",
            min_value=1, max_value=20, value=DEFAULT_MASTER_TAB, step=1,
        )

    submitted = st.form_submit_button("🚀 Start Processing", use_container_width=True)


# ─────────────────────────────────────────────────────────────
# PROCESSING
# ─────────────────────────────────────────────────────────────
if submitted:
    valid_urls = [u.strip() for u in sheet_urls if u.strip()]
    if not valid_urls:
        st.error("Please enter at least one Sheet URL.")
    else:
        st.info(f"Processing **{len(valid_urls)} sheet(s)** with **{workers} workers**...")

        progress_bar  = st.progress(0.0)
        status_text   = st.empty()

        def _progress(done: int, total: int):
            pct = done / total if total else 0
            progress_bar.progress(min(pct, 1.0))
            status_text.markdown(
                f"⏳ **{done} / {total}** batches done &nbsp;|&nbsp; "
                f"{pct * 100:.0f}%"
            )

        try:
            results = asyncio.run(
                run_sheets(
                    valid_urls,
                    workers=int(workers),
                    master_tab=int(master_tab),
                    progress_cb=_progress,
                )
            )
            progress_bar.progress(1.0)
            status_text.success(f"✅ Done! Processed **{len(results)} URLs** across {len(valid_urls)} sheet(s).")
            st.session_state["results"] = results
        except Exception as e:
            st.error(f"❌ Error during processing: {e}")
            st.exception(e)


# ─────────────────────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────────────────────
if "results" in st.session_state and st.session_state["results"]:
    results = st.session_state["results"]
    df      = pd.DataFrame(results)

    st.divider()
    st.subheader("📊 Results")

    # ── Summary metrics ──────────────────────────────────────
    indexed     = df[df["status"].str.contains("✅", na=False)]
    not_indexed = df[df["status"].str.contains("❌", na=False)]
    stale       = df[df["status"].str.contains("⚠️", na=False)]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total URLs",     len(df))
    m2.metric("✅ Indexed",     len(indexed))
    m3.metric("❌ Not Indexed", len(not_indexed))
    m4.metric("⚠️ Stale / Unknown", len(stale))

    st.divider()

    # ── Sheet filter (if multiple sheets) ────────────────────
    if df["sheet_label"].nunique() > 1:
        sheet_options = ["All Sheets"] + sorted(df["sheet_label"].unique().tolist())
        selected_sheet = st.selectbox("Filter by Sheet", sheet_options, key="sheet_filter")
        view_df = df if selected_sheet == "All Sheets" else df[df["sheet_label"] == selected_sheet]
    else:
        view_df = df

    # ── Three dropdowns ──────────────────────────────────────
    st.subheader("🔎 Browse by Status")
    drop_col1, drop_col2, drop_col3 = st.columns(3)

    with drop_col1:
        st.markdown("### ✅ Indexed")
        idx_urls = view_df[view_df["status"].str.contains("✅", na=False)]["url"].tolist()
        if idx_urls:
            st.selectbox(
                f"{len(idx_urls)} indexed URL(s)",
                idx_urls,
                key="drop_indexed",
            )
        else:
            st.info("No indexed URLs")

    with drop_col2:
        st.markdown("### ❌ Not Indexed")
        ni_urls = view_df[view_df["status"].str.contains("❌", na=False)]["url"].tolist()
        if ni_urls:
            st.selectbox(
                f"{len(ni_urls)} not-indexed URL(s)",
                ni_urls,
                key="drop_not_indexed",
            )
        else:
            st.info("No unindexed URLs")

    with drop_col3:
        st.markdown("### ⚠️ Stale / Unknown")
        stale_urls = view_df[view_df["status"].str.contains("⚠️", na=False)]["url"].tolist()
        if stale_urls:
            st.selectbox(
                f"{len(stale_urls)} stale/unknown URL(s)",
                stale_urls,
                key="drop_stale",
            )
        else:
            st.info("No stale/unknown URLs")

    # ── Full results table ───────────────────────────────────
    st.divider()
    st.subheader("📋 Full Results Table")

    status_filter = st.multiselect(
        "Filter by status",
        options=["✅ Indexed", "❌ Not Indexed", "⚠️ Stale / Unknown"],
        default=["✅ Indexed", "❌ Not Indexed", "⚠️ Stale / Unknown"],
        key="status_filter",
    )

    def _matches(s):
        if "✅ Indexed" in status_filter and "✅" in s:
            return True
        if "❌ Not Indexed" in status_filter and "❌" in s:
            return True
        if "⚠️ Stale / Unknown" in status_filter and "⚠️" in s:
            return True
        return False

    filtered_df = view_df[view_df["status"].apply(_matches)].copy()
    filtered_df = filtered_df.sort_values("row_num").reset_index(drop=True)
    display_df  = filtered_df[["row_num", "url", "status", "sheet_label", "batch", "checked_at"]]
    display_df.columns = ["#", "URL", "Status", "Sheet", "Batch", "Checked At"]

    st.dataframe(display_df, use_container_width=True, height=400)

    # ── Download buttons ─────────────────────────────────────
    st.divider()
    dl1, dl2 = st.columns(2)

    with dl1:
        csv_data = filtered_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download filtered CSV",
            csv_data,
            "results_filtered.csv",
            "text/csv",
            use_container_width=True,
        )

    with dl2:
        if os.path.exists(OUTPUT_FILE):
            with open(OUTPUT_FILE, "rb") as f:
                st.download_button(
                    "📥 Download full Excel (.xlsx)",
                    f.read(),
                    "results.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
