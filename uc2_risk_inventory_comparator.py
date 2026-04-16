"""
uc2_risk_inventory_comparator.py
────────────────────────────────────────────────────────────────────
USE CASE 2 — Risk Inventory Taxonomy Comparator (CRO View)
Golden example template for intern teaching purposes.

What this app does:
  • Loads a risk inventory (CSV or built-in sample data)
  • Lets the user drill into any taxonomy node (L1 → L2)
  • Calls an LLM to compare how each business division quantifies
    and manages that risk — surfacing similarities & differences
  • Presents results across three tabs:
      Tab 1 — Side-by-side comparison matrix (colour-coded diffs)
      Tab 2 — CRO-ready narrative + recommended actions
      Tab 3 — Drill-down Q&A chat

Setup:
  pip install streamlit openai pandas
  Add OPENAI_API_KEY to .streamlit/secrets.toml
  Run: streamlit run uc2_risk_inventory_comparator.py
────────────────────────────────────────────────────────────────────
"""

import json
import streamlit as st
import pandas as pd
from openai import OpenAI           # pip install openai


# ── CHANGE 1 of 4: model name ─────────────────────────────────────
# gpt-5.2 is available but the current OpenAI flagship is gpt-5.4.
# Swap the string below to upgrade at any time.
MODEL = "gpt-5.2"


# ── 0. Page config ────────────────────────────────────────────────

st.set_page_config(
    page_title="Risk Inventory Comparator",
    page_icon="🏦",
    layout="wide",
)


# ── 1. Sample inventory data ──────────────────────────────────────
#
# In production replace with pd.read_csv(...) or a DB query.
# Embedded here so the app runs out of the box for teaching.

SAMPLE_INVENTORY = [
    # ── Credit Risk / Counterparty Risk ──────────────────────────
    {
        "record_id": "RISK-001", "risk_category": "Credit Risk",
        "risk_subcategory": "Counterparty Risk",
        "business_division": "Investment Banking",
        "quant_method": "IMM / SA-CCR hybrid",
        "key_assumptions": "99% confidence, 10-day holding period, ISDA netting applied",
        "primary_metric": "PFE (Potential Future Exposure)",
        "secondary_metrics": "CVA, Stress EAD",
        "limit_type": "Counterparty credit limit (CCL)",
        "approval_body": "Risk Capital & Counterparty Committee (RCCC)",
        "risk_rating": "High",
    },
    {
        "record_id": "RISK-002", "risk_category": "Credit Risk",
        "risk_subcategory": "Counterparty Risk",
        "business_division": "Retail Banking",
        "quant_method": "SA-CCR (standardised)",
        "key_assumptions": "95% confidence, 1-day holding period, no netting recognised",
        "primary_metric": "EAD (Exposure at Default)",
        "secondary_metrics": "ECL Stage 2 flags",
        "limit_type": "Single-name exposure cap",
        "approval_body": "Retail Risk Credit Committee",
        "risk_rating": "Medium",
    },
    {
        "record_id": "RISK-003", "risk_category": "Credit Risk",
        "risk_subcategory": "Counterparty Risk",
        "business_division": "Treasury",
        "quant_method": "SA-CCR + bilateral stress testing",
        "key_assumptions": "99% confidence, 10-day holding period, ISDA netting applied",
        "primary_metric": "PFE + Stress EAD",
        "secondary_metrics": "Margin shortfall estimate",
        "limit_type": "ALCO-approved counterparty limit",
        "approval_body": "ALCO",
        "risk_rating": "High",
    },
    {
        "record_id": "RISK-004", "risk_category": "Credit Risk",
        "risk_subcategory": "Counterparty Risk",
        "business_division": "Wealth Management",
        "quant_method": "Simple notional-based exposure",
        "key_assumptions": "No statistical model, conservative notional haircut",
        "primary_metric": "Gross notional exposure",
        "secondary_metrics": "None",
        "limit_type": "Client-level credit line",
        "approval_body": "WM Risk team",
        "risk_rating": "Low",
    },
    # ── Market Risk / Interest Rate Risk ─────────────────────────
    {
        "record_id": "RISK-005", "risk_category": "Market Risk",
        "risk_subcategory": "Interest Rate Risk",
        "business_division": "Investment Banking",
        "quant_method": "Historical simulation VaR (2yr lookback)",
        "key_assumptions": "99% confidence, 10-day horizon, 500 scenarios",
        "primary_metric": "VaR ($M)",
        "secondary_metrics": "Stressed VaR, DV01, Convexity",
        "limit_type": "VaR limit + DV01 ladder limits",
        "approval_body": "Market Risk Committee",
        "risk_rating": "High",
    },
    {
        "record_id": "RISK-006", "risk_category": "Market Risk",
        "risk_subcategory": "Interest Rate Risk",
        "business_division": "Treasury",
        "quant_method": "EVE sensitivity + NII at risk",
        "key_assumptions": "±200bps parallel shock, 1yr NII horizon",
        "primary_metric": "NII at Risk ($M)",
        "secondary_metrics": "EVE change, ΔNII per 100bps",
        "limit_type": "EVE % of Tier 1 capital",
        "approval_body": "ALCO",
        "risk_rating": "High",
    },
    {
        "record_id": "RISK-007", "risk_category": "Market Risk",
        "risk_subcategory": "Interest Rate Risk",
        "business_division": "Retail Banking",
        "quant_method": "NII sensitivity (simple repricing gap)",
        "key_assumptions": "±100bps, behavioural assumptions on mortgage prepayments",
        "primary_metric": "NII at Risk ($M)",
        "secondary_metrics": "Repricing gap by bucket",
        "limit_type": "NII sensitivity board limit",
        "approval_body": "Retail ALCO",
        "risk_rating": "Medium",
    },
    # ── Operational Risk / Cyber Risk ────────────────────────────
    {
        "record_id": "RISK-008", "risk_category": "Operational Risk",
        "risk_subcategory": "Cyber Risk",
        "business_division": "Investment Banking",
        "quant_method": "Scenario-based AMA",
        "key_assumptions": "99.9% confidence, 1yr horizon, external loss data benchmarking",
        "primary_metric": "OpVar / Scenario loss estimate ($M)",
        "secondary_metrics": "Control effectiveness score",
        "limit_type": "Risk appetite statement (qualitative + quantitative)",
        "approval_body": "Operational Risk Committee",
        "risk_rating": "High",
    },
    {
        "record_id": "RISK-009", "risk_category": "Operational Risk",
        "risk_subcategory": "Cyber Risk",
        "business_division": "Retail Banking",
        "quant_method": "Standardised approach (TSA) + key risk indicators",
        "key_assumptions": "Frequency × severity model, KRI breach thresholds",
        "primary_metric": "KRI breach count + expected loss ($M)",
        "secondary_metrics": "Incident count, recovery time",
        "limit_type": "KRI-based escalation triggers",
        "approval_body": "Retail Ops Risk Forum",
        "risk_rating": "Medium",
    },
]


# ── 2. Taxonomy + division lists ──────────────────────────────────

TAXONOMY = {
    "Credit Risk":      ["Counterparty Risk", "Credit Concentration Risk",
                         "Settlement Risk", "Country Risk"],
    "Market Risk":      ["Interest Rate Risk", "FX Risk",
                         "Equity Risk", "Commodity Risk"],
    "Operational Risk": ["Cyber Risk", "Fraud Risk",
                         "Conduct Risk", "Third-Party Risk"],
    "Liquidity Risk":   ["Funding Liquidity Risk", "Market Liquidity Risk",
                         "Intraday Liquidity Risk"],
}

ALL_DIVISIONS = sorted({r["business_division"] for r in SAMPLE_INVENTORY})


# ── 3. Cached resources ───────────────────────────────────────────

@st.cache_resource
def get_client() -> OpenAI:
    # CHANGE 2 of 4: OpenAI client + secret key name
    return OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


@st.cache_data(show_spinner=False)
def load_inventory() -> pd.DataFrame:
    return pd.DataFrame(SAMPLE_INVENTORY)


# ── 4. Session state ──────────────────────────────────────────────

DEFAULTS = {
    "comparison":         None,
    "comparison_key":     "",
    "selected_divisions": ALL_DIVISIONS,
    "qa_messages":        [],
}
for key, val in DEFAULTS.items():
    st.session_state.setdefault(key, val)


# ── 5. LLM comparison functions ───────────────────────────────────

COMPARISON_SYSTEM_PROMPT = """
You are a Chief Risk Officer preparing a cross-divisional risk
framework review for the board.

Analyse how different business divisions quantify and manage the same
risk type. Return ONLY valid JSON — no markdown, no prose.

Required schema:
{
  "taxonomy_node": "string",
  "divisions_analysed": ["string"],
  "common_approach": "Paragraph describing what all divisions share.",
  "key_differences": [
    {
      "dimension": "string",
      "description": "string — what differs and between which divisions",
      "materiality": "High | Medium | Low"
    }
  ],
  "metric_alignment": "High | Medium | Low",
  "assumption_gaps": ["string"],
  "cro_summary": "2-3 sentence executive narrative for a board pack.",
  "recommended_actions": ["string"]
}

Rules:
- Cite division names and record IDs when noting differences.
- Flag High-materiality differences prominently.
- Never invent information not present in the input records.
""".strip()


@st.cache_data(show_spinner=False)
def run_comparison(records_json: str, taxonomy_node: str) -> dict:
    """
    LLM comparison call, cached by (records_json, taxonomy_node).

    TEACHING NOTE: Changing the division selection changes records_json,
    which busts the cache automatically — no manual invalidation needed.
    """
    client = get_client()

    # CHANGE 3 of 4: OpenAI chat.completions with system as first message
    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=4096,
        messages=[
            {"role": "system", "content": COMPARISON_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f"Taxonomy: {taxonomy_node}\n\n"
                    f"Inventory records:\n{records_json}"
                ),
            },
        ],
    )

    # CHANGE 4 of 4: extract text from choices[0].message.content
    return json.loads(response.choices[0].message.content)


def stream_chat(system_prompt: str, messages: list):
    """
    Generator for streaming OpenAI responses — pass to st.write_stream().
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


# ── 6. Sidebar ────────────────────────────────────────────────────

with st.sidebar:
    st.header("Taxonomy selector")
    l1 = st.selectbox("Risk category (L1)", list(TAXONOMY.keys()), key="tax_l1")
    l2 = st.selectbox("Risk sub-category (L2)", TAXONOMY[l1], key="tax_l2")
    taxonomy_node = f"{l1} → {l2}"
    st.caption(f"Node: **{taxonomy_node}**")

    st.divider()
    st.subheader("Business divisions")
    selected_divisions = st.multiselect(
        "Include in comparison",
        ALL_DIVISIONS,
        default=ALL_DIVISIONS,
        key="selected_divisions",
    )

    st.divider()
    run_analysis = st.button(
        "Run comparison analysis",
        disabled=len(selected_divisions) < 2,
        use_container_width=True,
        help="Select at least 2 divisions to compare.",
    )


# ── 7. Filter inventory ───────────────────────────────────────────

inventory_df = load_inventory()
filtered_df = inventory_df[
    (inventory_df["risk_category"]    == l1) &
    (inventory_df["risk_subcategory"] == l2) &
    (inventory_df["business_division"].isin(selected_divisions))
].copy()


# ── 8. Trigger LLM comparison ─────────────────────────────────────

if run_analysis:
    if filtered_df.empty:
        st.warning("No inventory records found for this selection.")
    else:
        with st.spinner("Analysing across divisions — this may take 10–20s…"):
            records_json = filtered_df.to_json(orient="records", indent=2)
            st.session_state.comparison = run_comparison(records_json, taxonomy_node)
            st.session_state.qa_messages = []


# ── 9. Landing state ──────────────────────────────────────────────

st.title("Risk Inventory Comparator")
st.caption(f"Selected: **{taxonomy_node}** · {len(filtered_df)} records")

if filtered_df.empty:
    st.info("No records found for this node and division selection.")
    st.stop()

if st.session_state.comparison is None:
    st.info("Configure the taxonomy node in the sidebar, then click **Run comparison analysis**.")
    st.subheader("Inventory records for this selection")
    st.dataframe(
        filtered_df[["record_id", "business_division", "quant_method",
                     "primary_metric", "risk_rating"]],
        use_container_width=True,
        hide_index=True,
    )
    st.stop()

comp = st.session_state.comparison


# ── 10. Header summary ────────────────────────────────────────────

alignment_icon = {"High": "🟢", "Medium": "🟠", "Low": "🔴"}.get(
    comp.get("metric_alignment", ""), "⚪"
)
col1, col2, col3 = st.columns(3)
col1.metric("Divisions compared",      len(comp.get("divisions_analysed", [])))
col2.metric("Key differences flagged", len(comp.get("key_differences", [])))
col3.metric("Metric alignment",        f"{alignment_icon} {comp.get('metric_alignment','—')}")
st.divider()


# ── 11. Tabs ──────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(
    ["Comparison matrix", "CRO narrative", "Drill-down Q&A"]
)


# ── Tab 1: Comparison matrix ──────────────────────────────────────

with tab1:
    st.subheader("Side-by-side comparison")

    display_cols = ["record_id", "business_division", "quant_method",
                    "key_assumptions", "primary_metric", "secondary_metrics",
                    "approval_body", "risk_rating"]

    def highlight_rating(val):
        return {
            "High":   "background-color:#fee2e2; color:#991b1b",
            "Medium": "background-color:#fef3c7; color:#92400e",
            "Low":    "background-color:#d1fae5; color:#065f46",
        }.get(val, "")

    st.dataframe(
        filtered_df[display_cols].style.applymap(
            highlight_rating, subset=["risk_rating"]
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Key differences (LLM-identified)")
    differences = comp.get("key_differences", [])
    if not differences:
        st.success("No material differences detected.")
    else:
        for diff in differences:
            mat  = diff.get("materiality", "Low")
            icon = {"High": "🔴", "Medium": "🟠", "Low": "🟢"}.get(mat, "⚪")
            with st.expander(f"{icon} **{diff.get('dimension','')}** — {mat} materiality"):
                st.write(diff.get("description", ""))

    st.subheader("What all divisions share")
    st.info(comp.get("common_approach", "—"))


# ── Tab 2: CRO narrative ──────────────────────────────────────────

with tab2:
    st.subheader("Executive summary (CRO narrative)")
    cro_text = comp.get("cro_summary", "—")
    st.markdown(
        f"""<div style="
            background:#f8fafc; border-left:4px solid #3b82f6;
            border-radius:6px; padding:16px 20px;
            font-size:15px; line-height:1.8; color:#1e293b;
        ">{cro_text}</div>""",
        unsafe_allow_html=True,
    )

    st.subheader("Assumption gaps")
    for gap in comp.get("assumption_gaps", []):
        st.warning(gap)

    st.subheader("Recommended actions")
    for i, action in enumerate(comp.get("recommended_actions", []), 1):
        st.markdown(f"**{i}.** {action}")

    st.divider()
    export_text = (
        f"RISK INVENTORY COMPARISON — {taxonomy_node}\n"
        f"{'=' * 60}\n\n"
        f"EXECUTIVE SUMMARY\n{cro_text}\n\n"
        "ASSUMPTION GAPS\n"
        + "\n".join(f"• {g}" for g in comp.get("assumption_gaps", [])) + "\n\n"
        "RECOMMENDED ACTIONS\n"
        + "\n".join(
            f"{i}. {a}"
            for i, a in enumerate(comp.get("recommended_actions", []), 1)
        )
    )
    st.download_button(
        "Download CRO summary (.txt)",
        data=export_text,
        file_name=(
            f"cro_summary_{l1.replace(' ','_')}_{l2.replace(' ','_')}.txt"
        ),
        mime="text/plain",
    )


# ── Tab 3: Drill-down Q&A ─────────────────────────────────────────

with tab3:
    st.subheader("Ask questions about this taxonomy node")

    def build_drilldown_prompt(df: pd.DataFrame, comp: dict) -> str:
        records_summary = df[
            ["record_id", "business_division", "quant_method",
             "key_assumptions", "primary_metric"]
        ].to_markdown(index=False)
        return (
            f"You are a senior risk analyst supporting a CRO review of:\n"
            f"Taxonomy: {taxonomy_node}\n\n"
            f"RAW INVENTORY RECORDS:\n{records_summary}\n\n"
            f"COMPARISON ANALYSIS:\n{json.dumps(comp, indent=2)}\n\n"
            "Answer questions about specific divisions, metrics, or "
            "assumptions. Be precise and cite record IDs where relevant."
        )

    for msg in st.session_state.qa_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if not st.session_state.qa_messages:
        st.caption("Suggested questions:")
        cols = st.columns(2)
        starters = [
            "Which division has the most conservative assumptions?",
            "Are the primary metrics consistent across divisions?",
            "What is the biggest governance gap?",
            "Which differences should be escalated to the board?",
        ]
        for col, question in zip(cols * 2, starters):
            if col.button(question, use_container_width=True):
                st.session_state.qa_messages.append(
                    {"role": "user", "content": question}
                )
                st.rerun()

    if prompt := st.chat_input("Ask about divisions, metrics, assumptions…"):
        st.session_state.qa_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            reply = st.write_stream(
                stream_chat(
                    system_prompt=build_drilldown_prompt(filtered_df, comp),
                    messages=st.session_state.qa_messages,
                )
            )

        st.session_state.qa_messages.append({"role": "assistant", "content": reply})
