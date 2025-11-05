import pandas as pd
import os

# ==========================================================
# 1️⃣ โหลดข้อมูลทั้งหมดของแต่ละจังหวัด
# ==========================================================
def load_all_sites(site_paths):
    """
    site_paths: dict ที่ mapping จังหวัด -> path ของไฟล์
    ต้องมีคอลัมน์ SITEID, LOCATION ID, PROVINCE_E
    """
    dfs = []
    for province, path in site_paths.items():
        if not os.path.exists(path):
            print(f"⚠️ ไฟล์ไม่พบ: {path}")
            continue
        df = pd.read_excel(path)
        # ปรับชื่อคอลัมน์ให้ตรงกัน
        df = df.rename(columns={'LOCATION ID': 'SITEID', 'PROVINCE_E': 'PROVINCE'})
        df['PROVINCE'] = province
        dfs.append(df[['SITEID', 'PROVINCE']])
    all_sites = pd.concat(dfs, ignore_index=True)
    return all_sites


# ==========================================================
# 2️⃣ รวม TT ที่ซ้ำกันภายใน site เดียวกัน (TT เดียว = นับครั้งเดียว)
# ==========================================================
def calculate_site_downtime_by_tt(df):
    """
    คำนวณ downtime ต่อ site โดย:
    - รวมข้อมูล TT ซ้ำ (TICKETID เดียวกันใน site เดียวกัน)
    - ใช้ค่า DOWN_TIME ที่มากที่สุดต่อ TICKETID
    - แปลงเป็นชั่วโมง
    """
    df = df.copy()
    df['DOWN_TIME'] = pd.to_numeric(df['DOWN_TIME'], errors='coerce').fillna(0)

    df_unique = (
        df.groupby(['SITE_7DIGITS', 'TICKETID'], as_index=False)['DOWN_TIME']
          .max()
    )

    site_downtime = (
        df_unique.groupby('SITE_7DIGITS', as_index=False)['DOWN_TIME']
          .sum()
          .rename(columns={'DOWN_TIME': 'DOWNTIME_MINS'})
    )
    site_downtime['DOWNTIME_HR'] = site_downtime['DOWNTIME_MINS'] / 60.0
    return site_downtime


# ==========================================================
# 3️⃣ Availability ต่อ Site
# ==========================================================
def calculate_service_availability_by_site(df, total_hours):
    """
    ใช้ downtime ต่อ site ที่นับเฉพาะ TT เดียวกันเพียงครั้งเดียว
    """
    site_dt = calculate_site_downtime_by_tt(df)
    site_dt['Availability (%)'] = ((total_hours - site_dt['DOWNTIME_HR']) / total_hours * 100).clip(lower=0)
    return site_dt


# ==========================================================
# 4️⃣ Availability ต่อ Province
# ==========================================================
def calculate_service_availability_by_province_from_site_dt(site_dt, site_df, total_hours):
    """
    site_dt: DataFrame จาก calculate_site_downtime_by_tt()
    site_df: DataFrame รวม site ทั้งหมด (SITEID, PROVINCE)
    """
    tmp = site_df.rename(columns={'SITEID':'SITE_7DIGITS'})
    merged = tmp.merge(site_dt, on='SITE_7DIGITS', how='left')
    merged['DOWNTIME_HR'] = merged['DOWNTIME_HR'].fillna(0)

    df_prov = merged.groupby('PROVINCE', as_index=False).agg({
        'SITE_7DIGITS': 'nunique',
        'DOWNTIME_HR': 'sum'
    }).rename(columns={'SITE_7DIGITS':'Site Count'})

    df_prov['Total Time (hr)'] = df_prov['Site Count'] * total_hours
    df_prov['Availability (%)'] = ((df_prov['Total Time (hr)'] - df_prov['DOWNTIME_HR']) / df_prov['Total Time (hr)'] * 100).round(4)
    return df_prov[['PROVINCE','Site Count','DOWNTIME_HR','Total Time (hr)','Availability (%)']]


# ==========================================================
# 5️⃣ แยก Availability เป็น Site ตามจังหวัด
# ==========================================================
def calculate_site_availability_by_province(site_dt, site_df):
    """
    รวม site availability ตามจังหวัด
    """
    tmp = site_df.rename(columns={'SITEID': 'SITE_7DIGITS'})
    merged = tmp.merge(site_dt, on='SITE_7DIGITS', how='left')
    merged['DOWNTIME_HR'] = merged['DOWNTIME_HR'].fillna(0)
    merged['Availability (%)'] = merged['Availability (%)'].fillna(100)
    return merged[['PROVINCE', 'SITE_7DIGITS', 'DOWNTIME_HR', 'Availability (%)']].sort_values(['PROVINCE','SITE_7DIGITS'])


# ==========================================================
# 6️⃣ Fault Rate, Fault Clear
# ==========================================================
def calculate_fault_rate(df, site_df):
    """ Fault Rate = Site Fault / All Site """
    site_faults = df['SITE_7DIGITS'].nunique()
    all_sites = site_df['SITEID'].nunique()
    return round((site_faults / all_sites) * 100, 2)

def calculate_fault_clear(df):
    """ Fault Clear = TT ที่ปิดได้ในเวลา / TT ทั้งหมด """
    df = df.copy()
    df['DOWN_TIME_HR'] = pd.to_numeric(df['DOWN_TIME'], errors='coerce').fillna(0) / 60.0
    df['TRUEURGENCY'] = pd.to_numeric(df['TRUEURGENCY'], errors='coerce').fillna(0)
    closed_on_time = df[df['DOWN_TIME_HR'] <= df['TRUEURGENCY']]
    ratio = len(closed_on_time) / len(df) if len(df) > 0 else 0
    return round(ratio * 100, 2)
