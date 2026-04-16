"""
uc1_risk_report_insights.py
────────────────────────────────────────────────────────────────────
USE CASE 1 — Monthly Risk Report Insight Extractor
Golden example template for intern teaching purposes.

What this app does:
  • Accepts up to 6 monthly risk report PDFs via sidebar upload
  • Calls an LLM to extract structured insights (JSON) from each report
  • Presents results across three tabs:
      Tab 1 — Risk appetite heatmap + key event timeline
      Tab 2 — Metric movements per risk type
      Tab 3 — Cross-report Q&A chat

Setup:
  pip install streamlit openai pypdf pandas
  Add OPENAI_API_KEY to .streamlit/secrets.toml
  Run: streamlit run uc1_risk_report_insights.py
────────────────────────────────────────────────────────────────────
"""

import json
import streamlit as st
import pandas as pd
from openai import OpenAI           # pip install openai
from io import BytesIO

try:
    from pypdf import PdfReader     # pip install pypdf
except ImportError:
    st.error("Run: pip install pypdf")
    st.stop()


# ── CHANGE 1 of 4: model name ─────────────────────────────────────
# gpt-5.2 is available but the current OpenAI flagship is gpt-5.4.
# Swap the string below to upgrade at any time.
MODEL = "gpt-5.2"


# ── 0. Page config ────────────────────────────────────────────────

st.set_page_config(
    page_title="Risk Report Insights",
    page_icon="📊",
    layout="wide",
)


# ── 1. Cached resources ───────────────────────────────────────────

@st.cache_resource
def get_client() -> OpenAI:
    # CHANGE 2 of 4: OpenAI client + secret key name
    return OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# ── 2. Session state initialisation ──────────────────────────────

DEFAULTS = {
    "reports":     [],
    "processing":  False,
    "risk_filter": "All",
    "div_filter":  "All",
    "qa_messages": [],
}
for key, val in DEFAULTS.items():
    st.session_state.setdefault(key, val)


# ── 3. LLM helpers ────────────────────────────────────────────────

EXTRACTION_SYSTEM_PROMPT = """
You are a senior risk analyst at a bank.
Extract structured data from a monthly risk report.
Return ONLY valid JSON — no markdown fences, no commentary.

Required schema (use null for missing fields, never invent numbers):
{
  "month": "YYYY-MM",
  "report_title": "string",
  "executive_summary": "1-2 sentence overview",
  "risk_appetite_status": {
    "Credit": "Within | Near limit | Breach",
    "Market": "Within | Near limit | Breach",
    "Operational": "Within | Near limit | Breach",
    "Liquidity": "Within | Near limit | Breach"
  },
  "key_events": [
    {
      "event_name": "string",
      "risk_type": "Credit | Market | Operational | Liquidity | Other",
      "business_unit": "string",
      "severity": "High | Medium | Low",
      "description": "Max 2 sentences.",
      "action_taken": "string"
    }
  ],
  "metric_movements": [
    {
      "risk_type": "string",
      "metric_name": "string",
      "value": 0.0,
      "unit": "$M | bps | % | x",
      "prior_value": 0.0,
      "business_division": "string",
      "direction": "increase | decrease | stable"
    }
  ]
}
""".strip()


def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(BytesIO(file_bytes))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


@st.cache_data(show_spinner=False)
def extract_report(report_text: str, month_label: str) -> dict:
    """Call the LLM once per report; result is cached by (text, month)."""
    client = get_client()

    # CHANGE 3 of 4: OpenAI uses client.chat.completions.create().
    # The system prompt is the first message with role="system".
    # (Anthropic uses a separate system= argument instead.)
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=4096,
        messages=[
            {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f"Month: {month_label}\n\n"
                    f"Report text (first 12 000 chars):\n{report_text[:12_000]}"
                ),
            },
        ],
    )

    # CHANGE 4 of 4: response text lives at choices[0].message.content
    # (Anthropic uses response.content[0].text instead.)
    return json.loads(response.choices[0].message.content)


def stream_chat(system_prompt: str, messages: list):
    """
    Generator that yields text tokens from a streaming OpenAI response.
    Pass directly to st.write_stream() for live token display.

    TEACHING NOTE: Anthropic exposes stream.text_stream as a ready-made
    generator. With OpenAI we build an equivalent generator manually by
    iterating over chunks and extracting chunk.choices[0].delta.content.
    """
    client = get_client()
    stream = client.chat.completions.create(
        model=MODEL,
        max_tokens=1024,
        stream=True,
        messages=[{"role": "system", "content": system_prompt}] + messages,
    )
    for chunk in stream:
        delta = chunk.choices[0].delta.content
        if delta:
            yield delta


# ── 4. Sidebar ────────────────────────────────────────────────────

with st.sidebar:
    st.header("Upload reports")
    uploaded_files = st.file_uploader(
        "Select up to 6 monthly PDFs",
        type="pdf",
        accept_multiple_files=True,
        help="Name files as YYYY-MM_report.pdf (e.g. 2025-01_report.pdf)",
    )
    run_extraction = st.button(
        "Extract insights",
        disabled=not uploaded_files,
        use_container_width=True,
    )
    st.divider()
    st.subheader("Filters")
    st.selectbox(
        "Risk type",
        ["All", "Credit", "Market", "Operational", "Liquidity"],
        key="risk_filter",
    )
    st.selectbox(
        "Business division",
        ["All", "Investment Banking", "Retail Banking",
         "Wealth Management", "Treasury"],
        key="div_filter",
    )


# ── 5. Extraction pipeline ────────────────────────────────────────

if run_extraction and uploaded_files:
    st.session_state.reports = []
    st.session_state.qa_messages = []
    progress = st.progress(0, text="Starting extraction…")

    for i, f in enumerate(uploaded_files):
        month_label = f.name[:7] if len(f.name) >= 7 else f"Month {i+1}"
        with st.spinner(f"Extracting {month_label}…"):
            raw_text  = extract_text_from_pdf(f.read())
            extracted = extract_report(raw_text, month_label)
            st.session_state.reports.append(extracted)
        progress.progress((i + 1) / len(uploaded_files), text=f"Done: {month_label}")

    progress.empty()
    st.rerun()


# ── 6. Guard: nothing extracted yet ──────────────────────────────

if not st.session_state.reports:
    st.title("Risk Report Insight Extractor")
    st.info("Upload up to 6 monthly PDFs in the sidebar, then click **Extract insights**.")
    st.stop()


# ── 7. Build DataFrames ───────────────────────────────────────────

def build_metrics_df(reports: list) -> pd.DataFrame:
    rows = []
    for r in reports:
        for m in r.get("metric_movements", []):
            rows.append({
                "month":       r["month"],
                "risk_type":   m.get("risk_type", ""),
                "metric_name": m.get("metric_name", ""),
                "value":       m.get("value"),
                "unit":        m.get("unit", ""),
                "prior_value": m.get("prior_value"),
                "division":    m.get("business_division", ""),
                "direction":   m.get("direction", ""),
            })
    return pd.DataFrame(rows)


def build_events_df(reports: list) -> pd.DataFrame:
    rows = []
    for r in reports:
        for e in r.get("key_events", []):
            rows.append({
                "month":         r["month"],
                "event_name":    e.get("event_name", ""),
                "risk_type":     e.get("risk_type", ""),
                "business_unit": e.get("business_unit", ""),
                "severity":      e.get("severity", ""),
                "description":   e.get("description", ""),
                "action_taken":  e.get("action_taken", ""),
            })
    return pd.DataFrame(rows)


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    if st.session_state.risk_filter != "All" and "risk_type" in df.columns:
        df = df[df["risk_type"] == st.session_state.risk_filter]
    if st.session_state.div_filter != "All":
        col = "division" if "division" in df.columns else "business_unit"
        if col in df.columns:
            df = df[df[col] == st.session_state.div_filter]
    return df


reports    = st.session_state.reports
metrics_df = apply_filters(build_metrics_df(reports))
events_df  = apply_filters(build_events_df(reports))


# ── 8. Header KPIs ────────────────────────────────────────────────

st.title("Risk Report Insights")
st.caption("Analysed " + str(len(reports)) + " report(s): "
           + ", ".join(r["month"] for r in reports))

high_count = (events_df["severity"] == "High").sum()  if not events_df.empty else 0
med_count  = (events_df["severity"] == "Medium").sum() if not events_df.empty else 0

col1, col2, col3 = st.columns(3)
col1.metric("Reports loaded",         len(reports))
col2.metric("High-severity events",   high_count)
col3.metric("Medium-severity events", med_count)
st.divider()


# ── 9. Tabs ───────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["Timeline & events", "Metric movements", "Q&A chat"])


# ── Tab 1: Heatmap + event list ───────────────────────────────────

with tab1:
    st.subheader("Risk appetite status — month × risk type")

    risk_types = ["Credit", "Market", "Operational", "Liquidity"]
    months     = sorted(r["month"] for r in reports)
    heatmap    = {rt: {} for rt in risk_types}

    for r in reports:
        status = r.get("risk_appetite_status", {})
        for rt in risk_types:
            heatmap[rt][r["month"]] = status.get(rt, "—")

    heatmap_df = pd.DataFrame(heatmap, index=months)

    def colour_cell(val):
        return {
            "Breach":     "background-color:#fee2e2; color:#991b1b",
            "Near limit": "background-color:#fef3c7; color:#92400e",
            "Within":     "background-color:#d1fae5; color:#065f46",
        }.get(val, "")

    st.dataframe(
        heatmap_df.style.applymap(colour_cell),
        use_container_width=True,
    )

    st.subheader("Key risk events")
    if events_df.empty:
        st.info("No events match the current filters.")
    else:
        for _, row in events_df.iterrows():
            icon = {"High": "🔴", "Medium": "🟠", "Low": "🟢"}.get(row["severity"], "⚪")
            with st.expander(
                f"{icon} {row['month']} · {row['event_name']} "
                f"— {row['risk_type']} / {row['business_unit']}"
            ):
                st.write(row["description"])
                if row["action_taken"]:
                    st.caption(f"**Action taken:** {row['action_taken']}")


# ── Tab 2: Metric movements ───────────────────────────────────────

with tab2:
    st.subheader("Metric movements across reports")
    if metrics_df.empty:
        st.info("No metrics match the current filters.")
    else:
        for metric_name, group in metrics_df.groupby("metric_name"):
            st.markdown(f"**{metric_name}**")
            display = group[["month", "risk_type", "division",
                              "value", "unit", "direction"]].copy()
            display["direction"] = display["direction"].map(
                {"increase": "▲", "decrease": "▼", "stable": "–"}
            ).fillna("–")
            st.dataframe(
                display.sort_values("month"),
                use_container_width=True,
                hide_index=True,
            )


# ── Tab 3: Q&A chat ───────────────────────────────────────────────

with tab3:
    st.subheader("Ask questions about the 6-month history")

    def build_qa_system_prompt(reports: list) -> str:
        sections = []
        for r in reports:
            high_events = [
                e["event_name"] for e in r.get("key_events", [])
                if e.get("severity") == "High"
            ]
            sections.append(
                f"=== {r['month']} ===\n"
                f"Summary: {r.get('executive_summary', 'N/A')}\n"
                f"Appetite: {r.get('risk_appetite_status', {})}\n"
                f"High events: {high_events}"
            )
        return (
            "You are a senior risk analyst. Answer questions about the "
            "6-month risk report history below. Cite specific months and "
            "figures. Flag trends spanning multiple months.\n\n"
            "REPORT HISTORY:\n" + "\n\n".join(sections)
        )

    for msg in st.session_state.qa_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Ask about events, metrics, or trends…"):
        st.session_state.qa_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            reply = st.write_stream(
                stream_chat(
                    system_prompt=build_qa_system_prompt(reports),
                    messages=st.session_state.qa_messages,
                )
            )

        st.session_state.qa_messages.append({"role": "assistant", "content": reply})
