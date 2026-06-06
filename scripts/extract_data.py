#!/usr/bin/env python3
"""
Phase 2: Data Extraction Script
Extracts all CSVs from Excel sources into /data/ folder.
Row indices are 0-based throughout (pandas convention).
"""
import pandas as pd
import shutil
from pathlib import Path

BASE = Path("/Users/vajonlim/Uni/FIT 2179 Data Visualisation /Assignment 2.1")
DATA_DIR = BASE / "Data"
OUT_DIR = BASE / "data"
ONLINE_DIR = DATA_DIR / "online"

OUT_DIR.mkdir(exist_ok=True)

print("=" * 60)
print("Phase 2: Data Extraction")
print("=" * 60)

# ─────────────────────────────────────────────────────────────────────────────
# 2.1  GP workforce by MM category (2020–2025)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.1] GP data by MM category …")

MM_META = {
    "MM1": ("Major cities",             1),
    "MM2": ("Large regional cities",    2),
    "MM3": ("Medium rural towns",       3),
    "MM4": ("Small rural towns",        4),
    "MM5": ("Very small rural towns",   5),
    "MM6": ("Remote communities",       6),
    "MM7": ("Very remote communities",  7),
}
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]

gp_file = DATA_DIR / "dataset-gp-calendar-years-2020-to-2025.xlsx"
rows = []
for mm, (label, order) in MM_META.items():
    df = pd.read_excel(gp_file, sheet_name=mm, header=None)
    for i, yr in enumerate(YEARS):
        col = i + 1   # columns 1–6 map to years 2020–2025
        rows.append({
            "mm_category": mm,
            "mm_label":    label,
            "mm_order":    order,
            "year":        yr,
            "total_gps":      pd.to_numeric(df.iloc[10, col], errors="coerce"),
            "single_mms_gps": pd.to_numeric(df.iloc[11, col], errors="coerce"),
            "gp_fte":         pd.to_numeric(df.iloc[14, col], errors="coerce"),
            "vr_gps_fte":     pd.to_numeric(df.iloc[17, col], errors="coerce"),
            "nonvr_gps_fte":  pd.to_numeric(df.iloc[18, col], errors="coerce"),
            "trainee_gps_fte":pd.to_numeric(df.iloc[19, col], errors="coerce"),
        })

gp_df = pd.DataFrame(rows)
gp_df.to_csv(OUT_DIR / "gp_by_mm.csv", index=False)
print(f"  ✓ gp_by_mm.csv  — {len(gp_df)} rows")
print(gp_df[gp_df["year"] == 2025][["mm_category", "total_gps", "gp_fte"]].to_string(index=False))

# ─────────────────────────────────────────────────────────────────────────────
# 2.2  Health workforce by remoteness
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.2] Workforce remoteness …")
wf = pd.read_excel(DATA_DIR / "Health_Workforce_Remoteness_Data_v2.xlsx", sheet_name=0)
wf.columns = [c.strip() for c in wf.columns.astype(str)]
print(f"  Raw columns: {list(wf.columns)}")

# Normalise column names
rename = {}
for c in wf.columns:
    cl = c.lower()
    if "remoteness" in cl or "area" in cl:       rename[c] = "remoteness"
    elif "profession" in cl or "occup" in cl:    rename[c] = "profession"
    elif "headcount" in cl:                      rename[c] = "headcount"
    elif "fte" in cl or "full" in cl:            rename[c] = "fte"
    elif "hour" in cl:                           rename[c] = "avg_hours"
    elif "age" in cl:                            rename[c] = "avg_age"
wf = wf.rename(columns=rename)
wf.to_csv(OUT_DIR / "workforce_remoteness.csv", index=False)
print(f"  ✓ workforce_remoteness.csv  — {len(wf)} rows, cols: {list(wf.columns)}")

# ─────────────────────────────────────────────────────────────────────────────
# 2.3  Health workforce by state
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.3] Workforce state …")
ws = pd.read_excel(DATA_DIR / "Health_Workforce_State_Data.xlsx", sheet_name=0)
ws.columns = [c.strip() for c in ws.columns.astype(str)]
ws.to_csv(OUT_DIR / "workforce_state.csv", index=False)
print(f"  ✓ workforce_state.csv  — {len(ws)} rows, cols: {list(ws.columns)}")

# ─────────────────────────────────────────────────────────────────────────────
# 2.4–2.6 + DC6/DC7  Patient Experience time-series
# Layout: row 5 = year headers, row 7 = "did not need" %, row 8 = "needed" %
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.4–2.6] Patient experience time-series …")

PEX = [
    ("DC2_PEX_2425_GPS_T4_to_6.xlsx",   "Table 4",  "GP",            "patient_exp_gp.csv"),
    ("DC3_PEX_2425_AHM_T7_to_9.xlsx",   "Table 7",  "After-Hours GP","patient_exp_afterhours.csv"),
    ("DC5_PEX_2425_DEN_T13_to_15.xlsx", "Table 13", "Dental",        "patient_exp_dental.csv"),
    ("DC6_PEX_2425_HSP_T16_to_18.xlsx", "Table 16", "Hospital",      "patient_exp_hospital.csv"),
    ("DC7_PEX_2425_EMG_T19_to_21.xlsx", "Table 19", "ED",            "patient_exp_ed.csv"),
]

for fname, sheet, service, out in PEX:
    fpath = DATA_DIR / "extracted" / fname
    try:
        df = pd.read_excel(fpath, sheet_name=sheet, header=None)
        year_labels = [str(v).strip() for v in df.iloc[5, 1:].dropna().values]
        did_not = pd.to_numeric(df.iloc[7, 1:len(year_labels)+1], errors="coerce").values
        needed  = pd.to_numeric(df.iloc[8, 1:len(year_labels)+1], errors="coerce").values
        ts = pd.DataFrame({
            "year":           year_labels,
            "service":        service,
            "did_not_need_pct": did_not,
            "needed_pct":       needed,
        }).dropna(subset=["needed_pct"])
        ts.to_csv(OUT_DIR / out, index=False)
        print(f"  ✓ {out}  — {len(ts)} rows  (latest: {ts.iloc[-1]['year']} {ts.iloc[-1]['needed_pct']:.1f}%)")
    except Exception as e:
        print(f"  ✗ {out}: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# 2.7  DC10 Mental Health — age breakdown (estimates in '000, compute %)
# Table 27.1: row 14 = "needed", row 22 = "did not need"; cols 1–7 = age groups
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.7] Mental health age breakdown …")

AGE_GROUPS_7 = ["15–24", "25–34", "35–44", "45–54", "55–64", "65–74", "75+"]

dc10_path = DATA_DIR / "extracted" / "DC10_PEX_2425_MHC_T27_to_28.xlsx"
mental_age_rows = []
try:
    df10 = pd.read_excel(dc10_path, sheet_name="Table 27.1", header=None)
    for j, ag in enumerate(AGE_GROUPS_7):
        col = j + 1
        n  = pd.to_numeric(df10.iloc[14, col], errors="coerce")
        nn = pd.to_numeric(df10.iloc[22, col], errors="coerce")
        if pd.notna(n) and pd.notna(nn) and (n + nn) > 0:
            mental_age_rows.append({
                "age_group":              ag,
                "service":                "Mental Health",
                "needed_estimate_000":    n,
                "did_not_need_est_000":   nn,
                "needed_pct":             round(n / (n + nn) * 100, 1),
            })
    mh_df = pd.DataFrame(mental_age_rows)
    mh_df.to_csv(OUT_DIR / "patient_exp_mental.csv", index=False)
    print(f"  ✓ patient_exp_mental.csv  — {len(mh_df)} rows")
    print(mh_df[["age_group","needed_pct"]].to_string(index=False))
except Exception as e:
    print(f"  ✗ Mental health: {e}")
    mh_df = pd.DataFrame()

# ─────────────────────────────────────────────────────────────────────────────
# 2.7b  Age breakdown chart data  (GP, Dental, Mental Health × age group)
# DC2 Table 5.1 / DC5 Table 14.1: row 8 = "did not need" ('000), row 9 = "needed" ('000)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.7b] Age breakdown for Chart 8 …")

AGE_GROUPS_8 = ["15–24", "25–34", "35–44", "45–54", "55–64", "65–74", "75–84", "85+"]

age_records = []

for svc, fname, sheet in [
    ("GP",     "DC2_PEX_2425_GPS_T4_to_6.xlsx",   "Table 5.1"),
    ("Dental", "DC5_PEX_2425_DEN_T13_to_15.xlsx", "Table 14.1"),
]:
    try:
        df_a = pd.read_excel(DATA_DIR / "extracted" / fname, sheet_name=sheet, header=None)
        for j, ag in enumerate(AGE_GROUPS_8):
            col = j + 1
            nn = pd.to_numeric(df_a.iloc[8, col], errors="coerce")
            n  = pd.to_numeric(df_a.iloc[9, col], errors="coerce")
            if pd.notna(n) and pd.notna(nn) and (n + nn) > 0:
                age_records.append({
                    "age_group":  ag,
                    "service":    svc,
                    "needed_pct": round(n / (n + nn) * 100, 1),
                })
        print(f"  ✓ {svc} age breakdown extracted")
    except Exception as e:
        print(f"  ✗ {svc} age breakdown: {e}")

# Add mental health (7 groups, harmonise to same structure)
for row in mental_age_rows:
    age_records.append({
        "age_group":  row["age_group"],
        "service":    "Mental Health",
        "needed_pct": row["needed_pct"],
    })

age_df = pd.DataFrame(age_records)
age_df.to_csv(OUT_DIR / "patient_exp_age.csv", index=False)
print(f"  ✓ patient_exp_age.csv  — {len(age_df)} rows")

# ─────────────────────────────────────────────────────────────────────────────
# 2.8  Heatmap CSV  (6 services × 5 representative years)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.8] Heatmap CSV …")

HEATMAP_YEARS = ["2013–14", "2016–17", "2019–20", "2022–23", "2024–25"]
# Match on first 4-digit year within the label
HEATMAP_FIRST = {y: y[:4] for y in HEATMAP_YEARS}

SVC_FILES = {
    "GP":            OUT_DIR / "patient_exp_gp.csv",
    "After-Hours GP":OUT_DIR / "patient_exp_afterhours.csv",
    "Dental":        OUT_DIR / "patient_exp_dental.csv",
    "Hospital":      OUT_DIR / "patient_exp_hospital.csv",
    "ED":            OUT_DIR / "patient_exp_ed.csv",
}

hm_rows = []
for svc, fpath in SVC_FILES.items():
    if not fpath.exists():
        print(f"  ⚠ Missing {fpath.name}")
        continue
    sdf = pd.read_csv(fpath)
    sdf["year"] = sdf["year"].astype(str).str.strip()
    for hy in HEATMAP_YEARS:
        key = hy[:4]
        match = sdf[sdf["year"].str.startswith(key)]
        if len(match):
            hm_rows.append({
                "year":       hy,
                "service":    svc,
                "needed_pct": float(match.iloc[0]["needed_pct"]),
            })

hm_df = pd.DataFrame(hm_rows)
hm_df.to_csv(OUT_DIR / "patient_exp_heatmap.csv", index=False)
print(f"  ✓ patient_exp_heatmap.csv  — {len(hm_df)} rows")
print(hm_df.pivot(index="year", columns="service", values="needed_pct").to_string())

# ─────────────────────────────────────────────────────────────────────────────
# 2.9  PHIDU Victoria
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.9] PHIDU Victoria …")

PHIDU = DATA_DIR / "phidu_data_pha_vic.xlsx"

def read_phidu_raw(sheet):
    """Read a PHIDU sheet with no headers (rows 0-4 are metadata)."""
    return pd.read_excel(PHIDU, sheet_name=sheet, header=None)

def phidu_pha_rows(df):
    """Filter to actual PHA rows: 5-digit integer codes only."""
    codes = pd.to_numeric(df.iloc[5:, 0], errors="coerce")
    mask = codes.notna() & codes.between(20000, 29999)
    return df.iloc[5:][mask.values].reset_index(drop=True)

try:
    # ── ED data (cols: 0=code, 1=name, 12=persons number, 13=ASR per 100k persons) ──
    ed_raw = read_phidu_raw("ED_total_sex")
    ed_rows = phidu_pha_rows(ed_raw)
    base = pd.DataFrame({
        "pha_code":        pd.to_numeric(ed_rows.iloc[:, 0], errors="coerce").astype(int),
        "pha_name":        ed_rows.iloc[:, 1].astype(str).str.strip(),
        "ed_rate_per_100k": pd.to_numeric(ed_rows.iloc[:, 13], errors="coerce"),
    })
    print(f"  ✓ ED data  — {len(base)} PHAs")

    # ── Self-assessed health (cols: 0=code, 1=name, 3=ASR per 100) ──
    try:
        sah_raw = read_phidu_raw("Estimates_self_assessed_health")
        sah_rows = phidu_pha_rows(sah_raw)
        sah_clean = pd.DataFrame({
            "pha_code":       pd.to_numeric(sah_rows.iloc[:, 0], errors="coerce").astype(int),
            "poor_health_pct": pd.to_numeric(sah_rows.iloc[:, 3], errors="coerce"),
        }).drop_duplicates(subset=["pha_code"])
        base = base.merge(sah_clean, on="pha_code", how="left")
        print(f"  ✓ Self-assessed health merged ({len(sah_clean)} rows)")
    except Exception as e:
        base["poor_health_pct"] = None
        print(f"  ⚠ SAH merge: {e}")

    # ── Population (cols: 0=code, 1=name, 3=total population) ──
    try:
        pop_raw = read_phidu_raw("Population_proportion")
        pop_rows = phidu_pha_rows(pop_raw)
        pop_clean = pd.DataFrame({
            "pha_code":   pd.to_numeric(pop_rows.iloc[:, 0], errors="coerce").astype(int),
            "population": pd.to_numeric(pop_rows.iloc[:, 3], errors="coerce"),
        }).drop_duplicates(subset=["pha_code"])
        pop_clean["pha_code"] = pop_clean["pha_code"].astype(int)
        base = base.merge(pop_clean, on="pha_code", how="left")
        print(f"  ✓ Merged population ({len(pop_clean)} rows)")
    except Exception as e:
        base["population"] = None
        print(f"  ⚠ Pop merge: {e}")

    # ── Area type classification (Metro vs Regional) ──
    # Greater Melbourne PHA codes are in range 20001–20999 (SA2-based PHA coding)
    base["area_type"] = base["pha_code"].apply(
        lambda c: "Metro Melbourne" if 20001 <= c <= 20999 else "Regional Victoria"
    )

    base.to_csv(OUT_DIR / "phidu_vic.csv", index=False)
    print(f"  ✓ phidu_vic.csv  — {len(base)} rows")
    print(f"     area_type counts: {base['area_type'].value_counts().to_dict()}")
    print(base[["pha_code","pha_name","ed_rate_per_100k","poor_health_pct","population"]].head(3).to_string(index=False))

except Exception as e:
    import traceback
    print(f"  ✗ PHIDU error: {e}")
    traceback.print_exc()

# ─────────────────────────────────────────────────────────────────────────────
# 2.10–2.15  Copy online CSVs into /data/
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2.10–2.15] Copying online data files …")

COPIES = [
    ("gp_visits_by_mm.csv",                    "gp_visits_by_mm.csv"),
    ("hospitalization_by_remoteness.csv",       "hospitalization_by_remoteness.csv"),
    ("mortality_by_remoteness.csv",             "mortality_by_remoteness.csv"),
    ("patient_experience_by_region_2024_25.csv","patient_experience_metro_vs_remote.csv"),
    ("access_barriers_by_remoteness.csv",       "access_barriers.csv"),
    ("workforce_per_100k_by_remoteness.csv",    "workforce_per_100k.csv"),
    ("population_by_remoteness.csv",            "population_by_remoteness.csv"),
    ("hero_key_stats.csv",                      "hero_key_stats.csv"),
]

for src_name, dst_name in COPIES:
    src = ONLINE_DIR / src_name
    dst = OUT_DIR / dst_name
    if src.exists():
        shutil.copy2(src, dst)
        print(f"  ✓ {dst_name}")
    else:
        print(f"  ✗ Missing: {src_name}")

# ─────────────────────────────────────────────────────────────────────────────
# Summary
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
csv_files = sorted(OUT_DIR.glob("*.csv"))
total_kb = sum(f.stat().st_size for f in csv_files) / 1024
print(f"Done!  {len(csv_files)} CSV files  •  {total_kb:.1f} KB total")
for f in csv_files:
    print(f"  {f.name:<45} {f.stat().st_size/1024:>6.1f} KB")
print("=" * 60)
