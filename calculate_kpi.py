import pandas as pd
from kpi_functions import *

def main():
    total_hours = 26 * 24  # ชั่วโมงในเดือนตุลาคม
    # path ของไฟล์ TT
    tt_file = "data/TT Oct.xlsx"

    # path ของไฟล์ site แต่ละจังหวัด
    site_paths = {
        "Chiang Mai": "data/site_list/CMI.xlsx",
        "Chiang Rai": "data/site_list/CRI.xlsx",
        "Kamphaeng Phet": "data/site_list/KPP.xlsx",
        "Lampang": "data/site_list/LPG.xlsx",
        "Lamphun": "data/site_list/LPN.xlsx",
        "Mae Hong Son": "data/site_list/MHS.xlsx",
        "Nan": "data/site_list/NAN.xlsx",
        "Phrae": "data/site_list/PCB.xlsx",
        "Phetchabun": "data/site_list/PCT.xlsx",
        "Phrae": "data/site_list/PHE.xlsx",
        "Phitsanulok": "data/site_list/PSN.xlsx",
        "Prachinburi": "data/site_list/PYO.xlsx",
        "Sukhothai": "data/site_list/SKT.xlsx",
        "Tak": "data/site_list/TAK.xlsx",
        "Uttaradit": "data/site_list/UTR.xlsx",
    }

     # ======================================================
    # 📥 LOAD DATA
    # ======================================================
    print("📂 กำลังโหลดข้อมูล Site ...")
    site_df = load_all_sites(site_paths)

    print("📂 กำลังโหลดข้อมูล TT ...")
    tt_df = pd.read_excel(tt_file)

    # ======================================================
    # 🧮 คำนวณ Availability
    # ======================================================
    print("⚙️ คำนวณ Availability ...")
    site_availability = calculate_service_availability_by_site(tt_df, total_hours)

    # รวม province เข้า site availability
    site_by_province = calculate_site_availability_by_province(site_availability, site_df)

    # ======================================================
    # 📊 รวมเป็น summary ต่อจังหวัด
    # ======================================================
    province_summary = calculate_service_availability_by_province_from_site_dt(
        site_availability, site_df, total_hours
    )

    # ======================================================
    # 💾 Export to Excel (multi-sheet)
    # ======================================================
    output_path = "KPI_Site_Availability_Report.xlsx"
    print(f"💾 บันทึกไฟล์ไปยัง: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # รวมทั้งหมด
        site_by_province.to_excel(writer, sheet_name='All_Provinces', index=False)

        # แยกตามจังหวัด
        for province in sorted(site_by_province['PROVINCE'].unique()):
            df_sub = site_by_province[site_by_province['PROVINCE'] == province]
            df_sub.to_excel(writer, sheet_name=province[:31], index=False)

        # สรุป Province
        province_summary.to_excel(writer, sheet_name='Summary_By_Province', index=False)

    print("✅ สร้างไฟล์ KPI_Site_Availability_Report.xlsx เรียบร้อยแล้ว!")

if __name__ == "__main__":
    main()