from __future__ import annotations

import importlib.util
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import streamlit as st


def ensure_dependency(import_name: str, pip_requirement: str) -> None:
    """
    ตรวจสอบแพ็กเกจที่จำเป็น

    หากไม่มีใน Python Environment ปัจจุบัน
    ระบบจะพยายามติดตั้งให้อัตโนมัติ
    """
    if importlib.util.find_spec(import_name) is not None:
        return

    try:
        subprocess.check_call(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                pip_requirement,
            ]
        )
        importlib.invalidate_caches()

    except Exception as exc:
        st.error(
            f"ไม่พบแพ็กเกจ {import_name} "
            f"และระบบติดตั้งอัตโนมัติไม่สำเร็จ: {exc}"
        )

        st.info("กรุณารันคำสั่งต่อไปนี้ใน Terminal")

        st.code(
            f'"{sys.executable}" -m pip install "{pip_requirement}"',
            language="powershell",
        )

        st.stop()


# ตรวจสอบ XlsxWriter ก่อน Import processor.py
ensure_dependency(
    import_name="xlsxwriter",
    pip_requirement="XlsxWriter>=3.2,<4.0",
)

from processor import ProcessingError, build_output


APP_DIR = Path(__file__).resolve().parent

MASTER_PATH = (
    APP_DIR
    / "masters"
    / "accessory_master.xlsx"
)

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")


st.set_page_config(
    page_title="AMS Checklist & AllNow",
    page_icon="📄",
    layout="centered",
)


st.title("AMS Checklist & AllNow")

st.caption(
    "อัปโหลดไฟล์ AMS เพื่อสร้าง Checklist "
    "แยกตามเลขที่เอกสาร และ AllNow รวม"
)


# ตรวจสอบไฟล์ Master
if not MASTER_PATH.exists():
    st.error(
        "ไม่พบไฟล์ Master: "
        f"{MASTER_PATH.relative_to(APP_DIR)}"
    )
    st.stop()


# Session State
if "generated_excel" not in st.session_state:
    st.session_state.generated_excel = None

if "generated_name" not in st.session_state:
    st.session_state.generated_name = None

if "generated_summary" not in st.session_state:
    st.session_state.generated_summary = None


# Upload AMS
uploaded_file = st.file_uploader(
    label="ไฟล์ AMS",
    type=["xls", "xlsx"],
    help="รองรับไฟล์ที่ดาวน์โหลดจากระบบ AMS โดยตรง",
)


if uploaded_file is not None:
    st.write(f"ไฟล์ที่เลือก: `{uploaded_file.name}`")

    generate_button = st.button(
        "สร้าง Checklist และ AllNow",
        type="primary",
        use_container_width=True,
    )

    if generate_button:
        try:
            # ล้างผลลัพธ์เดิม
            st.session_state.generated_excel = None
            st.session_state.generated_name = None
            st.session_state.generated_summary = None

            with st.spinner(
                "กำลังอ่านข้อมูล AMS, Filter และสร้างไฟล์..."
            ):
                result = build_output(
                    ams_content=uploaded_file.getvalue(),
                    ams_filename=uploaded_file.name,
                    master_path=MASTER_PATH,
                )

            branch_text = (
                "-".join(result.branch_codes)
                if result.branch_codes
                else "UNKNOWN"
            )

            timestamp = datetime.now(
                BANGKOK_TZ
            ).strftime("%Y%m%d_%H%M%S")

            output_name = (
                f"AMS_Checklist_AllNow_"
                f"{branch_text}_"
                f"{timestamp}.xlsx"
            )

            st.session_state.generated_excel = result.content
            st.session_state.generated_name = output_name
            st.session_state.generated_summary = result

        except ProcessingError as exc:
            st.session_state.generated_excel = None
            st.error(str(exc))

        except Exception as exc:
            st.session_state.generated_excel = None
            st.exception(exc)


# แสดงผลลัพธ์
if st.session_state.generated_excel is not None:
    result = st.session_state.generated_summary

    st.success("สร้างไฟล์สำเร็จ")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        label="ข้อมูล AMS",
        value=f"{result.source_rows:,} แถว",
    )

    col2.metric(
        label="หลัง Filter",
        value=f"{result.filtered_rows:,} แถว",
    )

    col3.metric(
        label="Checklist",
        value=f"{result.document_count:,} Sheet",
    )

    st.write(
        f"Match Master สำเร็จ "
        f"**{result.matched_rows:,}** จาก "
        f"**{result.filtered_rows:,}** รายการ"
    )

    # แสดงรายการที่ไม่พบใน Master
    if result.unmatched_items:
        st.warning(
            "พบชื่ออุปกรณ์ที่ไม่มีใน Master "
            "ระบบปล่อยช่องอุปกรณ์ส่วนควบว่าง "
            "โดยไม่เดาข้อมูล"
        )

        with st.expander(
            f"รายการที่ไม่พบใน Master "
            f"({len(result.unmatched_items)})"
        ):
            st.dataframe(
                {
                    "รายการทรัพย์สิน":
                        result.unmatched_items
                },
                use_container_width=True,
                hide_index=True,
            )

    # Download
    st.download_button(
        label="ดาวน์โหลดไฟล์ Excel",
        data=st.session_state.generated_excel,
        file_name=st.session_state.generated_name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )