import streamlit as st
import tempfile
from pathlib import Path

st.set_page_config(page_title="think-cell Updater (22 charts)", page_icon="📊", layout="centered")
st.title("📊 think-cell — Update 22 charts (download a new PPT)")

st.markdown(
    """
Upload your PowerPoint and Excel. The app updates **multiple think-cell elements** **in place**  
using fixed Excel ranges, then gives you a **downloadable updated copy**.

**Requirements**
- Windows + local **Microsoft Office (Excel + PowerPoint)** + **think-cell**
- Python package: `pywin32`
- Uses think-cell API: `UpdateChart(presentation, name, range, transposed)`
- Updating **breaks the live Excel link** for each updated element (think-cell behavior)
"""
)

# ===================== ONE-TIME CONFIG — EDIT JUST THIS ===========================
# Provide 22 mappings (or any number). "tc_name" must match the element's UpdateChart Name in PPT.
# "address" can be A1-style (e.g., "A1:D6") OR a workbook named range (e.g., "RevenueBlock").
# Set "transposed": True only if your Excel orientation is rotated vs think-cell datasheet.
MAPPINGS = [
    {"tc_name": "Process_Eff",   "sheet": "Sheet1", "address": "A2:Z12",    "transposed": True},
    {"tc_name": "Process_Cost","sheet": "Sheet1", "address": "A16:Z26",    "transposed": True},
    {"tc_name": "FPA_Eff",       "sheet": "Sheet1", "address": "A30:Z35",   "transposed": True},
    {"tc_name": "FPA_Cost",       "sheet": "Sheet1", "address": "A39:Z44",   "transposed": True},
    {"tc_name": "RTR_Eff",       "sheet": "Sheet1", "address": "A160:Z164",   "transposed": True},
    {"tc_name": "RTR_Cost",       "sheet": "Sheet1", "address": "A168:Z172", "transposed": True},
    {"tc_name": "ITC-NC_Eff",       "sheet": "Sheet1", "address": "A106:Z111",  "transposed": True},
    {"tc_name": "ITC-NC_Cost",       "sheet": "Sheet1", "address": "A115:Z120",  "transposed": True},
    {"tc_name": "ITC-C_Eff",       "sheet": "Sheet1", "address": "A86:Z92",  "transposed": True},
    {"tc_name": "ITC-C_Cost",       "sheet": "Sheet1", "address": "A96:Z102",  "transposed": True},
    {"tc_name": "PTP_Eff",       "sheet": "Sheet1", "address": "A142:Z147",  "transposed": True},
    {"tc_name": "PTP_Cost",       "sheet": "Sheet1", "address": "A151:Z156","transposed": True},
    {"tc_name": "Tax_Eff",       "sheet": "Sheet1", "address": "A176:Z180",  "transposed": True},
    {"tc_name": "Tax_Cost",       "sheet": "Sheet1", "address": "A184:Z188",  "transposed": True},
    {"tc_name": "Treasury_Eff",       "sheet": "Sheet1", "address": "A192:Z198",  "transposed": True},
    {"tc_name": "Treasury_Cost",       "sheet": "Sheet1", "address": "A202:Z208",  "transposed": True},
    {"tc_name": "Payroll_Eff",       "sheet": "Sheet1", "address": "A124:Z129",  "transposed": True},
    {"tc_name": "Payroll_Cost",       "sheet": "Sheet1", "address": "A133:Z138","transposed": True},
    {"tc_name": "IR_Eff",       "sheet": "Sheet1", "address": "A64:Z71",  "transposed": True},
    {"tc_name": "IR_Cost",       "sheet": "Sheet1", "address": "A75:Z82",  "transposed": True},
    {"tc_name": "Audit_Eff",       "sheet": "Sheet1", "address": "A48:Z52",  "transposed": True},
    {"tc_name": "Audit_Cost",       "sheet": "Sheet1", "address": "A56:Z60",  "transposed": True},
]
# ================================================================================

st.caption("Edit the MAPPINGS list in the script to match your 22 charts and ranges.")

# ---- Uploads (always required in this “download-a-copy” app)
col1, col2 = st.columns(2)
with col1:
    ppt_file = st.file_uploader("PowerPoint (.pptx/.pptm)", type=["pptx", "pptm"])
with col2:
    xlsx_file = st.file_uploader("Excel (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

out_basename = st.text_input("Output filename", value="Deck_UPDATED.pptx")
run = st.button("🚀 Update all & get download")

def _persist_upload(upload, suffix):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(upload.read()); tmp.flush(); tmp.close()
    return tmp.name

if run:
    if not ppt_file or not xlsx_file:
        st.error("Please upload both PowerPoint and Excel.")
        st.stop()

    # Dependencies
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as e:
        st.error(
            "Requires Windows with Office + think-cell + `pywin32`.\n"
            "Install: `pip install pywin32`.\n\n"
            f"Details: {e}"
        )
        st.stop()

    # Save uploads for Office to open
    ppt_path = _persist_upload(ppt_file, Path(ppt_file.name).suffix)
    xlsx_path = _persist_upload(xlsx_file, Path(xlsx_file.name).suffix)
    output_path = str(Path(tempfile.gettempdir()) / out_basename)

    xl = wb = pp = pres = None

    try:
        st.info("Starting Office automation…")

        # Initialize COM (STA)
        pythoncom.CoInitialize()

        # Excel (hidden)
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        # PowerPoint (keep visible to avoid “cannot hide” error)
        pp = win32.DispatchEx("PowerPoint.Application")
        pp.Visible = True
        try:
            pp.WindowState = 2  # minimize (optional)
        except Exception:
            pass

        # Open files
        wb = xl.Workbooks.Open(xlsx_path)
        pres = pp.Presentations.Open(FileName=ppt_path, WithWindow=False)

        # think-cell automation via Excel COM add-in
        try:
            tc = xl.COMAddIns("thinkcell.addin").Object
        except Exception:
            tc = None
        if tc is None:
            raise RuntimeError(
                "Could not access think-cell via Excel COM Add-Ins. "
                "Excel → File → Options → Add-ins → Manage: COM Add-ins → Go… → check 'think-cell'."
            )

        # Update each mapping
        results = []
        updated = 0
        for m in MAPPINGS:
            name = m["tc_name"]; sheet = m["sheet"]; addr = m["address"]; trans = bool(m["transposed"])
            try:
                ws = wb.Worksheets(sheet)
            except Exception as e:
                results.append((name, f"❌ Worksheet '{sheet}' not found: {e}"))
                continue
            try:
                rng = ws.Range(addr)
            except Exception:
                # fallback to workbook-level named range
                try:
                    rng = wb.Range(addr)
                except Exception as e2:
                    results.append((name, f"❌ Range '{addr}' not found on '{sheet}' (or as named range): {e2}"))
                    continue

            try:
                tc.UpdateChart(pres, name, rng, trans)
                results.append((name, f"✅ Updated from {sheet}!{addr}, transposed={trans}"))
                updated += 1
            except Exception as e:
                results.append((name, f"❌ UpdateChart failed: {e}"))

        # Save to a NEW file and offer as download
        pres.SaveAs(output_path)

        st.success(f"Updated {updated} of {len(MAPPINGS)} chart(s). See details below.")
        for name, msg in results:
            st.write(f"- **{name}** — {msg}")

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download updated PPT",
                data=f.read(),
                file_name=Path(output_path).name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        st.caption("Note: `UpdateChart` replaces each chart's datasheet and **breaks any live Excel link** for that element.")

    except Exception as e:
        st.error(f"Update failed: {e}")

    finally:
        # Cleanup COM objects safely
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if pp is not None:
                pp.Quit()
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
