"""
Grade Verification - web app (Streamlit)
Upload an approved REFERENCE spec and the SPEC to check (same grade-rule
category). Returns a colour-coded Excel: green = grading matches the reference,
red = mismatch.

Privacy: both PDFs are processed in memory and never saved to the server.
"""
import io
import gc
import streamlit as st

from grade_verify import verify_to_bytes

st.set_page_config(page_title="Grade Verification", page_icon="✅", layout="centered")

# Hide Streamlit chrome (menu, deploy/fork, footer, header)
st.markdown(
    """
    <style>
      [data-testid="stHeader"] {display: none !important;}
      [data-testid="stToolbar"] {display: none !important;}
      [data-testid="stToolbarActions"] {display: none !important;}
      [data-testid="stAppDeployButton"] {display: none !important;}
      [data-testid="stMainMenu"] {display: none !important;}
      #MainMenu {visibility: hidden !important;}
      header {visibility: hidden !important;}
      footer {visibility: hidden !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("✅ Grade Verification")
st.write(
    "Upload an **approved reference** spec and the **spec to check** (same "
    "grade-rule category). The app compares the size-to-size grade increments: "
    "**green** where the spec grades exactly like the reference, **red** where it differs."
)
st.info("Both PDFs are processed in memory and are **never saved to the server**.", icon="🔒")

if "uid" not in st.session_state:
    st.session_state.uid = 0

col1, col2 = st.columns(2)
with col1:
    ref = st.file_uploader("Reference (approved) PDF", type=["pdf"],
                           key=f"ref_{st.session_state.uid}")
with col2:
    spec = st.file_uploader("Spec to check PDF", type=["pdf"],
                            key=f"spec_{st.session_state.uid}")

tol = st.number_input("Match tolerance (cm)", min_value=0.0, value=0.0, step=0.1,
                      help="0.0 = increments must match exactly. Raise it to allow small rounding differences.")

if ref is not None and spec is not None and st.button("Verify grading", type="primary"):
    rbuf = io.BytesIO(ref.getvalue())
    sbuf = io.BytesIO(spec.getvalue())
    try:
        with st.spinner("Reading both PDFs and comparing grade increments..."):
            data, summary = verify_to_bytes(rbuf, sbuf, tol=tol)
        st.session_state.xlsx = data
        st.session_state.summary = summary
        st.session_state.fname = spec.name.rsplit(".", 1)[0] + "_grade_check.xlsx"
    except Exception as e:
        st.error(f"Could not verify these PDFs: {e}")
    finally:
        rbuf.close(); sbuf.close()
        del rbuf, sbuf
        gc.collect()

if "xlsx" in st.session_state:
    s = st.session_state.summary
    if not s["grade_match"]:
        st.warning(
            f"Grade rule templates differ — reference is “{s['ref_grade']}”, "
            f"spec is “{s['spec_grade']}”. These may not be the same category.",
            icon="⚠️",
        )
    c1, c2, c3 = st.columns(3)
    c1.metric("Measurements checked", s["checked"])
    c2.metric("Mismatched breaks", s["mismatch_breaks"])
    c3.metric("Measurements with issues", s["mismatch_rows"])
    if s["mismatch_breaks"] == 0:
        st.success("All shared measurements grade exactly like the reference.")
    else:
        st.error("Grading differences found — see the red cells and the Exceptions sheet.")
    if s["only_spec"] or s["only_ref"]:
        st.caption(f"Not verified: {s['only_spec']} measurement(s) only in the spec, "
                   f"{s['only_ref']} only in the reference.")

    st.download_button(
        "⬇️ Download grade-check Excel",
        data=st.session_state.xlsx,
        file_name=st.session_state.fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if st.button("Clear files from this session"):
        for k in ("xlsx", "summary", "fname"):
            st.session_state.pop(k, None)
        st.session_state.uid += 1
        gc.collect()
        st.rerun()

st.divider()
st.caption("Compares grade increments, not absolute sizes, so two different garments in the "
           "same grade-rule template verify correctly. Works on PDFs with real text tables.")
