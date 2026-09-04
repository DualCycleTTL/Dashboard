import io
import re
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Font

st.set_page_config(page_title="Dashboard Analisis Dual Cycle", layout="wide")

# ================================================================
# KONSTANTA DEFAULT (samakan dengan VBA)
# ================================================================
AMBANG_COMBO_MENIT_DEFAULT = 40
AMBANG_DUAL_MENIT_DEFAULT = 240  # = 4 jam (bukan 6 jam spt komentar VBA lama)
AMBANG_TWINLIFT_MENIT_DEFAULT = 1  # jarak DISC_LOAD_TS antar 2 kontainer combo

# Ukuran kontainer eligible utk Combo di VBA di-hardcode "= 20" (bukan
# pengaturan yang bisa diubah user), jadi di sini juga dikunci tetap 20ft
# supaya hasil selalu identik dengan VBA.
SIZE_ELIGIBLE = 20


# ================================================================
# HELPER - baca & siapkan data
# ================================================================

def klasifikasi_activity(val: object) -> str:
    s = str(val).upper()
    return "LOAD" if "LOAD" in s else "DISC"


_VBA_VAL_RE = re.compile(r"^\s*[+-]?\d+(\.\d+)?")


def vba_val(x: object) -> float:
    """
    Replikasi fungsi Val() di VBA: baca angka dari AWAL string sampai
    ketemu karakter non-angka pertama, sisanya diabaikan (mis. "20FT"
    -> 20, "  -5.5kg" -> -5.5). Kalau tidak ada angka sama sekali -> 0.
    Ini supaya CTR_SIZE yang formatnya campur teks (mis. "20FT") tetap
    dibaca sama seperti macro VBA aslinya, bukan malah jadi 0 seperti
    kalau pakai pd.to_numeric biasa.
    """
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    m = _VBA_VAL_RE.match(str(x))
    return float(m.group()) if m else 0.0


@st.cache_data(show_spinner=False, max_entries=3, ttl=1800)
def baca_file(file_bytes: bytes, filename: str, sheet_name=None):
    if filename.lower().endswith(".csv"):
        return {"__csv__": pd.read_csv(BytesIO(file_bytes))}
    xls = pd.ExcelFile(BytesIO(file_bytes))
    if sheet_name is not None:
        return {sheet_name: xls.parse(sheet_name)}
    return {name: xls.parse(name) for name in xls.sheet_names}


def bersihkan_ves_id(series: pd.Series) -> pd.Series:
    """
    Normalisasi kolom VES_ID jadi string biasa (dtype object), bukan
    dtype 'string'/ArrowDtype bawaan pandas versi baru. Ini PENTING
    karena .astype(str) pada kolom ber-dtype 'string' TIDAK mengubah
    nilai NaN/NA jadi teks "nan" -- NaN tetap float, sehingga tercampur
    dengan string lain dan bikin sorted()/unique() gagal dengan
    TypeError: '<' not supported between instances of 'float' and 'str'.

    Nilai kosong/NaN diseragamkan jadi pd.NA (bukan string "nan"),
    dan whitespace di depan/belakang dibuang (spasi ekor sering muncul
    di data VES_ID mentah, mis. "HUHE010     ").
    """
    s = series.astype("object")
    s = s.where(~pd.isna(s), pd.NA)  # tandai semua bentuk kosong (NaN/None/NA) secara seragam
    s = s.map(lambda v: v.strip() if isinstance(v, str) else v)
    s = s.map(lambda v: pd.NA if isinstance(v, str) and v == "" else v)
    return s


def siapkan_data(raw: pd.DataFrame, col_map: dict, size_eligible: int) -> pd.DataFrame:
    df = pd.DataFrame()
    df["VES_ID"] = bersihkan_ves_id(raw[col_map["ves_id"]])
    df["CTR_SIZE"] = raw[col_map["size"]].apply(vba_val)
    df["CAR_CHE_ID"] = raw[col_map["truck"]].astype(str).str.strip()
    df["ACTIVITY"] = raw[col_map["activity"]].apply(klasifikasi_activity)
    df["TS_G"] = pd.to_datetime(raw[col_map["ts_g"]], errors="coerce")
    df["TS_H"] = pd.to_datetime(raw[col_map["ts_h"]], errors="coerce")

    # kalau G kosong pakai H, kalau H kosong pakai G (spt VBA)
    both_invalid = df["TS_G"].isna() & df["TS_H"].isna()
    df["TS_G"] = df["TS_G"].fillna(df["TS_H"])
    df["TS_H"] = df["TS_H"].fillna(df["TS_G"])

    # ------------------------------------------------------------
    # PENTING: VBA TIDAK PERNAH MEMBUANG BARIS, walau kedua kolom
    # timestamp kosong. Baris begitu tetap diproses sbg 1 event
    # tersendiri dengan tanggal dummy CDate(-1) = 29 Des 1899.
    # Versi sebelumnya di app ini MEMBUANG baris tsb (dropna),
    # sehingga Total Event & persentase jadi beda dari hasil VBA
    # (mis. 62,5% vs 63,2%). Sekarang disamakan persis dgn VBA:
    # baris tetap dihitung, bukan dibuang.
    # ------------------------------------------------------------
    if both_invalid.any():
        dummy_ts = pd.Timestamp("1899-12-29")  # = CDate(-1) di VBA
        df.loc[both_invalid, "TS_G"] = dummy_ts
        df.loc[both_invalid, "TS_H"] = dummy_ts
        st.warning(
            f"{int(both_invalid.sum())} baris punya kedua kolom timestamp "
            f"(DISC_LOAD_TS & STACK_UNSTACK_TS) kosong/tidak valid. Mengikuti "
            f"perilaku VBA, baris ini TETAP diproses sbg event tersendiri "
            f"(tanggal dummy 29 Des 1899), bukan dibuang, supaya total & "
            f"persentase persis sama dengan hasil macro VBA."
        )

    # ------------------------------------------------------------
    # VES_ID kosong: baris TETAP diproses (tidak dibuang) supaya
    # Total Event/Dual Cycle/Combo tetap identik dengan VBA -- VES_ID
    # tidak dipakai sama sekali di layer Combo/Dual/Twinlift, cuma
    # dipakai belakangan di tab "Per Vessel". Baris begini ditandai
    # placeholder "(VES_ID Kosong)" supaya:
    #   1) tidak bikin error TypeError saat sorting di tab Per Vessel
    #      (lihat bersihkan_ves_id di atas kenapa NaN mentah bermasalah),
    #   2) tetap kelihatan/terlacak sbg baris data yang datanya kurang
    #      lengkap, bukan hilang diam-diam atau nyasar gabung ke kapal lain.
    # ------------------------------------------------------------
    ves_kosong = df["VES_ID"].isna()
    if ves_kosong.any():
        st.warning(
            f"{int(ves_kosong.sum())} baris punya VES_ID kosong. Baris ini tetap "
            f"diproses di analisis Dual Cycle/Twinlift (tidak dibuang), tapi di tab "
            f"'Per Vessel' dikelompokkan terpisah sbg \"(VES_ID Kosong)\"."
        )
        df.loc[ves_kosong, "VES_ID"] = "(VES_ID Kosong)"

    df = df.reset_index(drop=True)
    df["ROW_IDX"] = df.index
    return df


# ================================================================
# LAYER 1 - COMBO / SINGLE
# ================================================================

def layer1_combo(df: pd.DataFrame, ambang_combo: float, size_eligible: int) -> pd.DataFrame:
    n = len(df)
    ts_g = df["TS_G"].to_numpy()
    ts_h = df["TS_H"].to_numpy()
    size = df["CTR_SIZE"].to_numpy()
    activity = df["ACTIVITY"].to_numpy()

    assigned = np.zeros(n, dtype=bool)
    group_id = np.zeros(n, dtype=int)

    pairs = []
    truck_positions = df.groupby("CAR_CHE_ID").indices  # dict truk -> array posisi (0..n-1)

    thr_delta = np.timedelta64(int(round(ambang_combo * 60)), "s")

    for _, pos in truck_positions.items():
        for act in ("LOAD", "DISC"):
            idx_act = np.array([p for p in pos if activity[p] == act and size[p] == size_eligible])
            m = len(idx_act)
            if m < 2:
                continue

            # ------------------------------------------------------
            # min(gapG, gapH) <= ambang  <=>  gapG <= ambang ATAU
            # gapH <= ambang. Jadi kandidat cukup dicari lewat 2
            # sliding-window (satu berdasar TS_G, satu TS_H) yang
            # jauh lebih cepat (O(m log m)) daripada cek SEMUA
            # pasangan (O(m^2)) -- hasil akhirnya identik.
            # ------------------------------------------------------
            found = set()

            for ts_arr in (ts_g, ts_h):
                order = idx_act[np.argsort(ts_arr[idx_act])]
                sorted_ts = ts_arr[order]
                left = 0
                for right in range(len(order)):
                    while sorted_ts[right] - sorted_ts[left] > thr_delta:
                        left += 1
                    for b in range(left, right):
                        i, k = order[b], order[right]
                        if i > k:
                            i, k = k, i
                        found.add((int(i), int(k)))

            for i, k in found:
                gap_g = abs((ts_g[k] - ts_g[i]) / np.timedelta64(1, "m"))
                gap_h = abs((ts_h[k] - ts_h[i]) / np.timedelta64(1, "m"))
                gap = min(gap_g, gap_h)
                pairs.append((i, k, gap))

    # urutkan gap terkecil dulu, lalu greedy matching
    pairs.sort(key=lambda x: (x[2], x[0], x[1]))

    nxt = 0
    for i, k, _gap in pairs:
        if not assigned[i] and not assigned[k]:
            nxt += 1
            group_id[i] = nxt
            group_id[k] = nxt
            assigned[i] = assigned[k] = True

    for i in range(n):
        if not assigned[i]:
            nxt += 1
            group_id[i] = nxt
            assigned[i] = True

    out = df.copy()
    out["GROUP_ID"] = group_id
    return out


# ================================================================
# LAYER 1b - TWINLIFT (sub-analisis di dalam grup Combo)
# ================================================================

def deteksi_twinlift(df_combo: pd.DataFrame, ambang_twinlift: float, size_eligible: int):
    """
    Twinlift adalah kondisi khusus DI DALAM grup Combo (2 baris jadi 1
    event, lihat layer1_combo) yang memenuhi SEMUA syarat berikut:
      1) Muatan Combo 20ft -- otomatis terpenuhi krn Combo di layer1
         memang hanya dibentuk dari baris dengan CTR_SIZE == size_eligible
         (20ft), tapi tetap dicek ulang di sini supaya aman/eksplisit.
      2) Berasal dari kapal yang sama -> VES_ID kedua baris identik.
      3) Waktu DISC_LOAD_TS (TS_G) antar kedua baris berdekatan, yaitu
         selisihnya <= ambang_twinlift menit (default 1 menit).

    Grup Single (cuma 1 anggota, tidak ber-Combo) otomatis BUKAN
    Twinlift, ditandai "-" (bukan "Bukan Twinlift") supaya gampang
    dibedakan dari Combo yang gagal syarat Twinlift.

    Mengembalikan 2 dict:
      - status_map: {GROUP_ID: "Twinlift" / "Bukan Twinlift" / "-"}
      - gap_map:    {GROUP_ID: selisih DISC_LOAD_TS dalam menit (float),
                     None utk grup Single} -- disediakan supaya hasil
                     bisa diaudit/dicek manual oleh user, bukan cuma
                     dipercaya begitu saja sbg "black box".
    """
    status_map = {}
    gap_map = {}
    for gid, g in df_combo.groupby("GROUP_ID"):
        if len(g) != 2:
            status_map[gid] = "-"
            gap_map[gid] = None
            continue

        g = g.sort_values("ROW_IDX")
        r1, r2 = g.iloc[0], g.iloc[1]

        syarat_size = (r1["CTR_SIZE"] == size_eligible) and (r2["CTR_SIZE"] == size_eligible)
        syarat_kapal = r1["VES_ID"] == r2["VES_ID"]
        gap_disc_load = abs((r2["TS_G"] - r1["TS_G"]) / np.timedelta64(1, "m"))
        syarat_waktu = gap_disc_load <= ambang_twinlift

        gap_map[gid] = round(float(gap_disc_load), 2)
        if syarat_size and syarat_kapal and syarat_waktu:
            status_map[gid] = "Twinlift"
        else:
            status_map[gid] = "Bukan Twinlift"

    return status_map, gap_map


# ================================================================
# BENTUK EVENT
# ================================================================

def bentuk_event(df: pd.DataFrame) -> pd.DataFrame:
    # Vektorisasi penuh (bukan loop per grup) supaya tetap cepat
    # meski jumlah grup/event sangat banyak.
    is_disc = df["ACTIVITY"].to_numpy() == "DISC"
    evt_start = np.where(is_disc, df["TS_G"].to_numpy(), df["TS_H"].to_numpy())
    evt_end = np.where(is_disc, df["TS_H"].to_numpy(), df["TS_G"].to_numpy())

    tmp = pd.DataFrame(
        {
            "GROUP_ID": df["GROUP_ID"].to_numpy(),
            "ACTIVITY": df["ACTIVITY"].to_numpy(),
            "CAR_CHE_ID": df["CAR_CHE_ID"].to_numpy(),
            "EVT_START": evt_start,
            "EVT_END": evt_end,
        }
    )

    # PENTING: groupby harus sort=True (urut menaik berdasarkan GROUP_ID),
    # BUKAN sort=False (yang ikut urutan kemunculan baris di data asli).
    # Ini supaya "posisi ke-0,1,2,..." event di sini persis sama dengan
    # urutan "g = 1 To nGrp" di VBA -- karena posisi ini dipakai sbg
    # tie-breaker saat sorting kandidat pasangan Dual Cycle yang gap-nya
    # SAMA PERSIS (sering terjadi krn banyak event overlap = gap 0 menit).
    # Kalau urutannya beda dari VBA, greedy matching-nya bisa memilih
    # pasangan yang berbeda meski total event-nya tetap sama -- inilah
    # yang menyebabkan Dual Cycle vs Non Dual beda antara VBA & dashboard
    # padahal Total Event-nya sudah identik.
    events = tmp.groupby("GROUP_ID", sort=True).agg(
        ACTIVITY=("ACTIVITY", "first"),
        CAR_CHE_ID=("CAR_CHE_ID", "first"),
        START_TS=("EVT_START", "min"),
        END_TS=("EVT_END", "max"),
        N_ANGGOTA=("GROUP_ID", "size"),
    ).reset_index()

    events["CONTAINER_STATUS"] = np.where(events["N_ANGGOTA"] >= 2, "Combo", "Single")
    events = events.drop(columns=["N_ANGGOTA"])
    return events


# ================================================================
# LAYER 2 - DUAL CYCLE
# ================================================================

def layer2_dual(events: pd.DataFrame, ambang_dual: float) -> pd.DataFrame:
    events = events.reset_index(drop=True)
    n = len(events)
    start = events["START_TS"].to_numpy()
    end = events["END_TS"].to_numpy()
    activity = events["ACTIVITY"].to_numpy()

    assigned = np.zeros(n, dtype=bool)
    status = np.array(["Non Dual"] * n, dtype=object)

    pairs = []
    truck_positions = events.groupby("CAR_CHE_ID").indices

    for _, pos in truck_positions.items():
        pos_sorted = sorted(pos, key=lambda p: start[p])
        m = len(pos_sorted)
        for a in range(m - 1):
            i = pos_sorted[a]
            for b in range(a + 1, m):
                k = pos_sorted[b]
                gap_ab = (start[k] - end[i]) / np.timedelta64(1, "m")
                gap_ba = (start[i] - end[k]) / np.timedelta64(1, "m")
                if gap_ab >= 0:
                    gap = gap_ab
                elif gap_ba >= 0:
                    gap = gap_ba
                else:
                    gap = 0.0  # overlap waktu -> dianggap berdekatan

                if gap > ambang_dual:
                    # waktu mulai naik monoton -> aman berhenti di sini
                    break

                if activity[i] != activity[k]:
                    pairs.append((i, k, gap))

    pairs.sort(key=lambda x: (x[2], x[0], x[1]))
    for i, k, _gap in pairs:
        if not assigned[i] and not assigned[k]:
            status[i] = "Dual Cycle"
            status[k] = "Dual Cycle"
            assigned[i] = assigned[k] = True

    out = events.copy()
    out["STATUS"] = status
    return out


# ================================================================
# PENOMORAN EVENT_ID GLOBAL (urut truk sesuai kemunculan pertama
# di data asli, di dalam truk diurutkan berdasarkan START_TS)
# ================================================================

def beri_event_id(events: pd.DataFrame, df_asli: pd.DataFrame):
    truck_order = list(dict.fromkeys(df_asli["CAR_CHE_ID"].tolist()))
    rank = {tk: i for i, tk in enumerate(truck_order)}

    events = events.copy()
    events["_truck_rank"] = events["CAR_CHE_ID"].map(rank)
    events = events.sort_values(["_truck_rank", "START_TS"]).reset_index(drop=True)
    events["EVENT_ID"] = events.index + 1
    events = events.drop(columns=["_truck_rank"])

    event_id_map = dict(zip(events["GROUP_ID"], events["EVENT_ID"]))
    return events, event_id_map


def gabungkan_hasil(df: pd.DataFrame, events: pd.DataFrame, event_id_map: dict) -> pd.DataFrame:
    status_map = events.set_index("GROUP_ID")["STATUS"].to_dict()
    container_map = events.set_index("GROUP_ID")["CONTAINER_STATUS"].to_dict()
    twinlift_map = events.set_index("GROUP_ID")["TWINLIFT_STATUS"].to_dict()
    twinlift_gap_map = events.set_index("GROUP_ID")["TWINLIFT_GAP_MENIT"].to_dict()

    out = df.copy()
    out["EVENT_ID"] = out["GROUP_ID"].map(event_id_map)
    out["CONTAINER_STATUS"] = out["GROUP_ID"].map(container_map)
    out["STATUS"] = out["GROUP_ID"].map(status_map)
    out["TWINLIFT_STATUS"] = out["GROUP_ID"].map(twinlift_map)
    out["TWINLIFT_GAP_MENIT"] = out["GROUP_ID"].map(twinlift_gap_map)
    out = out.drop(columns=["GROUP_ID", "ROW_IDX"])
    return out


# ================================================================
# RINGKASAN / STATISTIK
# ================================================================

def hitung_ringkasan(events: pd.DataFrame, out_df: pd.DataFrame) -> dict:
    total_event = len(events)
    total_dual = int((events["STATUS"] == "Dual Cycle").sum())
    total_single = total_event - total_dual

    combo_dual = int(((events["CONTAINER_STATUS"] == "Combo") & (events["STATUS"] == "Dual Cycle")).sum())
    combo_single = int(((events["CONTAINER_STATUS"] == "Combo") & (events["STATUS"] == "Non Dual")).sum())
    single_dual = int(((events["CONTAINER_STATUS"] == "Single") & (events["STATUS"] == "Dual Cycle")).sum())
    single_single = int(((events["CONTAINER_STATUS"] == "Single") & (events["STATUS"] == "Non Dual")).sum())

    # ------------------------------------------------------
    # TWINLIFT -- sub-analisis di dalam grup Combo. Basisnya event
    # (bukan baris), sama seperti Combo/Single & Dual Cycle/Non Dual.
    # ------------------------------------------------------
    total_combo = int((events["CONTAINER_STATUS"] == "Combo").sum())
    total_twinlift = int((events["TWINLIFT_STATUS"] == "Twinlift").sum())
    total_combo_bukan_twinlift = total_combo - total_twinlift
    # Basis breakdown Twinlift = TOTAL EVENT (bukan Total Combo). "Bukan Twinlift"
    # di sini artinya semua event selain Twinlift, termasuk Single & Combo yang
    # bukan Twinlift.
    total_non_twinlift = total_event - total_twinlift
    pct_twinlift_of_total = (total_twinlift / total_event) if total_event else 0
    pct_non_twinlift_of_total = (total_non_twinlift / total_event) if total_event else 0
    pct_twinlift_of_combo = (total_twinlift / total_combo) if total_combo else 0

    dual_load = int(((out_df["STATUS"] == "Dual Cycle") & (out_df["ACTIVITY"] == "LOAD")).sum())
    dual_disc = int(((out_df["STATUS"] == "Dual Cycle") & (out_df["ACTIVITY"] == "DISC")).sum())
    single_load = int(((out_df["STATUS"] == "Non Dual") & (out_df["ACTIVITY"] == "LOAD")).sum())
    single_disc = int(((out_df["STATUS"] == "Non Dual") & (out_df["ACTIVITY"] == "DISC")).sum())

    container_load = dual_load + single_load
    container_disc = dual_disc + single_disc
    container_total = len(out_df)

    ev = events.copy()
    ev["BULAN"] = ev["START_TS"].dt.to_period("M")

    # ------------------------------------------------------
    # Breakdown bulanan -- 3 view terpisah, semuanya per event/ritase:
    #   1) Dual Cycle vs Non Dual (basis persis spt VBA)
    #   2) Combo vs Single (basis CONTAINER_STATUS, independen dari
    #      status Dual Cycle -- ini tambahan di luar VBA, sesuai
    #      permintaan, supaya kelihatan komposisi Combo/Single tiap
    #      bulan juga)
    #   3) Twinlift vs Bukan Twinlift (sub-analisis di dalam Combo,
    #      tambahan sesuai permintaan)
    # Semuanya disediakan sbg jumlah (dipakai internal) & persen
    # (dipakai buat ditampilkan, krn user minta lihat persen).
    # ------------------------------------------------------
    monthly = ev.groupby("BULAN").agg(
        total_event=("STATUS", "count"),
        dual=("STATUS", lambda s: int((s == "Dual Cycle").sum())),
        combo=("CONTAINER_STATUS", lambda s: int((s == "Combo").sum())),
        twinlift=("TWINLIFT_STATUS", lambda s: int((s == "Twinlift").sum())),
    )
    monthly["non_dual"] = monthly["total_event"] - monthly["dual"]
    monthly["single"] = monthly["total_event"] - monthly["combo"]
    monthly["combo_bukan_twinlift"] = monthly["combo"] - monthly["twinlift"]
    # Basis breakdown Twinlift = TOTAL EVENT bulan tsb (bukan Total Combo bulan tsb)
    monthly["non_twinlift"] = monthly["total_event"] - monthly["twinlift"]

    monthly["pct_dual"] = np.where(monthly["total_event"] > 0, monthly["dual"] / monthly["total_event"], 0)
    monthly["pct_non_dual"] = np.where(monthly["total_event"] > 0, monthly["non_dual"] / monthly["total_event"], 0)
    monthly["pct_combo"] = np.where(monthly["total_event"] > 0, monthly["combo"] / monthly["total_event"], 0)
    monthly["pct_single"] = np.where(monthly["total_event"] > 0, monthly["single"] / monthly["total_event"], 0)
    monthly["pct_twinlift"] = np.where(monthly["total_event"] > 0, monthly["twinlift"] / monthly["total_event"], 0)
    monthly["pct_non_twinlift"] = np.where(
        monthly["total_event"] > 0, monthly["non_twinlift"] / monthly["total_event"], 0
    )
    monthly["pct_twinlift_of_combo"] = np.where(
        monthly["combo"] > 0, monthly["twinlift"] / monthly["combo"], 0
    )
    monthly["pct_combo_bukan_twinlift_of_combo"] = np.where(
        monthly["combo"] > 0, monthly["combo_bukan_twinlift"] / monthly["combo"], 0
    )

    monthly = monthly.sort_index()
    monthly.index = monthly.index.astype(str)

    return {
        "total_event": total_event,
        "total_dual": total_dual,
        "total_single": total_single,
        "pct_dual": (total_dual / total_event) if total_event else 0,
        "combo_dual": combo_dual,
        "combo_single": combo_single,
        "single_dual": single_dual,
        "single_single": single_single,
        "total_combo": total_combo,
        "total_twinlift": total_twinlift,
        "total_combo_bukan_twinlift": total_combo_bukan_twinlift,
        "total_non_twinlift": total_non_twinlift,
        "pct_twinlift_of_total": pct_twinlift_of_total,
        "pct_non_twinlift_of_total": pct_non_twinlift_of_total,
        "pct_twinlift_of_combo": pct_twinlift_of_combo,
        "dual_load": dual_load,
        "dual_disc": dual_disc,
        "single_load": single_load,
        "single_disc": single_disc,
        "container_load": container_load,
        "container_disc": container_disc,
        "container_total": container_total,
        "monthly": monthly,
    }


# ================================================================
# EXPORT EXCEL (DATA SAJA -- versi ringan)
# ================================================================
#
# CATATAN PERFORMA: versi sebelumnya di sini membangun Excel dengan
# sheet Ringkasan penuh + beberapa PieChart/BarChart openpyxl + tabel
# breakdown bulanan, semuanya ditulis cell-per-cell. Itu berat & lambat
# terutama utk data besar (ratusan ribu baris), dan bikin proses
# "Download Hasil" jadi lelet. Sesuai permintaan, itu semua DIHAPUS.
# Yang tersisa cuma sheet "Data" mentah, ditulis pakai pandas
# ExcelWriter (jauh lebih ringan/cepat drpd loop manual + chart).
# Kalau butuh ringkasan/chart, olah sendiri dari data mentah ini di
# Excel/BI tool pilihan masing-masing.
# ================================================================

def build_excel_data_only(out_df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
    bio.seek(0)
    return bio.getvalue()


# ================================================================
# ================================================================
# UI STREAMLIT
# ================================================================
# ================================================================

# ----------------------------------------------------------------
# Path logo dihitung RELATIF terhadap lokasi file app.py ini sendiri
# (bukan relatif terhadap current working directory), supaya tetap
# ketemu walau di-deploy di Streamlit Cloud yang CWD-nya bisa beda
# dari folder script. File logo harus ada di:
#   <folder app.py>/assets/logo_caca_icon.png  (ikon CACA saja, tanpa teks)
#   <folder app.py>/assets/logo_pelindo.png
# ----------------------------------------------------------------
ASSETS_DIR = Path(__file__).resolve().parent / "assets"
LOGO_CACA_ICON_PATH = ASSETS_DIR / "caca.png"
LOGO_PATH = ASSETS_DIR / "logo_pelindo.png"


def _img_to_base64(path: Path) -> str:
    import base64
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def _render_header(icon_path: Path, brand_path: Path):
    """
    Header custom pakai flexbox HTML (bukan st.columns) supaya ikon,
    judul "CACA", subjudul, dan logo Pelindo di kanan bisa presisi
    sejajar vertikal (align-items: center) mirip layout referensi.
    """
    icon_ok = icon_path.exists()
    brand_ok = brand_path.exists()

    icon_html = (
        f'<img src="data:image/png;base64,{_img_to_base64(icon_path)}" '
        f'style="height:64px;width:auto;display:block;" />'
        if icon_ok else ""
    )
    brand_html = (
        f'<img src="data:image/png;base64,{_img_to_base64(brand_path)}" '
        f'style="height:64px;width:auto;display:block;" />'
        if brand_ok else ""
    )

    st.markdown(
        f"""
        <div style="display:flex;align-items:center;justify-content:space-between;
                    padding:6px 0 2px 0;">
            <div style="display:flex;align-items:center;gap:12px;">
                {icon_html}
                <div style="line-height:1.15;">
                    <div style="font-size:1.55rem;font-weight:800;color:#16324f;
                                letter-spacing:0.5px;">CACA</div>
                    <div style="font-size:0.72rem;color:#6b7280;">
                        Cycle Analysis and Cargo Optimalization
                    </div>
                </div>
            </div>
            <div>{brand_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not icon_ok:
        st.caption(f"⚠️ Logo CACA tidak ditemukan di `{icon_path}`.")
    if not brand_ok:
        st.caption(f"⚠️ Logo Pelindo tidak ditemukan di `{brand_path}`.")


_render_header(LOGO_CACA_ICON_PATH, LOGO_PATH)
st.divider()

# ----------------------------------------------------------------
# PENGATURAN -- ditaruh langsung di dashboard (bukan sidebar) biar
# layout terasa lebih lega/luas. Ukuran kontainer eligible utk Combo
# TIDAK lagi jadi pengaturan yang bisa diubah -- dikunci 20ft persis
# spt VBA (lihat SIZE_ELIGIBLE di atas), karena mengubahnya bisa bikin
# hasil beda dari VBA.
# ----------------------------------------------------------------
size_eligible = SIZE_ELIGIBLE

with st.expander("⚙️ Pengaturan Ambang Batas Analisis", expanded=False):
    pc1, pc2, pc3 = st.columns(3)
    with pc1:
        ambang_combo = st.number_input(
            "Ambang Combo (menit)", min_value=1, value=AMBANG_COMBO_MENIT_DEFAULT, step=5,
            help="Jarak waktu maksimum antar 2 baris size 20ft, truk & aktivitas sama, supaya dianggap 'Combo'.",
        )
    with pc2:
        ambang_dual = st.number_input(
            "Ambang Dual Cycle (menit)", min_value=1, value=AMBANG_DUAL_MENIT_DEFAULT, step=10,
            help="Jarak waktu maksimum antar 2 event beda aktivitas (LOAD vs DISC) dalam truk yang sama.",
        )
    with pc3:
        ambang_twinlift = st.number_input(
            "Ambang Twinlift (menit)", min_value=1, value=AMBANG_TWINLIFT_MENIT_DEFAULT, step=1,
            help=(
                "Jarak waktu maksimum DISC_LOAD_TS antar 2 kontainer dalam 1 Combo 20ft, "
                "yang berasal dari kapal (VES_ID) & truk yang sama, supaya dianggap 'Twinlift'."
            ),
        )
    st.caption(f"Ukuran kontainer eligible untuk Combo/Twinlift dikunci **{SIZE_ELIGIBLE} ft**, mengikuti macro VBA.")

st.subheader("📁 Upload Data")
uploaded = st.file_uploader("Upload file .xlsx atau .csv", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.session_state.pop("hasil", None)
    st.session_state.pop("_last_file_sig", None)
    st.info("Silakan upload file data untuk memulai analisis.")
    st.stop()

file_bytes = uploaded.getvalue()

# --------------------------------------------------------------
# Deteksi kalau file yang diupload BERBEDA dari sebelumnya
# (nama, ukuran, atau isinya beda) -> otomatis bersihkan hasil
# analisis lama supaya tidak nyangkut / bikin bingung. Ini juga
# yang bikin ganti-ganti file jadi tidak perlu reboot app lagi.
# --------------------------------------------------------------
file_sig = (uploaded.name, len(file_bytes), hash(file_bytes[:1_000_000]))
if st.session_state.get("_last_file_sig") != file_sig:
    st.session_state.pop("hasil", None)
    st.session_state["_last_file_sig"] = file_sig

try:
    sheets = baca_file(file_bytes, uploaded.name)
except Exception as e:
    st.error(
        "❌ Gagal membaca file yang diupload. Pastikan file tidak corrupt dan "
        "formatnya benar-benar .xlsx / .xls / .csv."
    )
    with st.expander("Detail error (untuk dilaporkan)"):
        st.exception(e)
    st.stop()

if not sheets:
    st.error("❌ File tidak berisi sheet/data apa pun.")
    st.stop()

sheet_name = None
if len(sheets) > 1:
    sheet_name = st.selectbox("Pilih sheet data", list(sheets.keys()))
else:
    sheet_name = list(sheets.keys())[0]

raw = sheets[sheet_name]
st.caption(f"File terbaca: {len(raw)} baris data.")

# ----------------------------------------------------------------
# Pemetaan kolom TIDAK ditampilkan lagi sbg langkah terpisah -- kolom
# dideteksi otomatis di background pakai guess() (cocokkan nama kolom
# yg mengandung kata kunci), supaya alurnya cukup: upload -> klik
# jalankan -> langsung tampil dashboard, tanpa perlu isi form kolom.
# ----------------------------------------------------------------
cols = list(raw.columns)


def guess(options, keywords, default_idx=0):
    for kw in keywords:
        for i, c in enumerate(options):
            if kw.lower() in str(c).lower():
                return i
    return default_idx


col_map = {
    "ves_id": cols[guess(cols, ["ves"])],
    "size": cols[guess(cols, ["size", "ctr_size"])],
    "truck": cols[guess(cols, ["car_che", "truck", "che"])],
    "activity": cols[guess(cols, ["activity", "aktivitas"])],
    "ts_g": cols[guess(cols, ["disc_load", "disc_loading"])],
    "ts_h": cols[guess(cols, ["stack_unstack", "unstack_stack"])],
}

run = st.button("▶️ Jalankan Analisis Dual Cycle", type="primary")

# ----------------------------------------------------------------
# Kalau ada hasil analisis lama tersimpan di session (dari sebelum
# kode di-update, mis. sebelum field ringkasan baru ditambahkan),
# buang saja cache lama itu supaya tidak KeyError -- minta user
# klik jalankan lagi. Cukup cek 1 key penanda skema paling baru.
# ----------------------------------------------------------------
if not run and "hasil" in st.session_state:
    _cached_summary = st.session_state["hasil"].get("summary", {})
    if "total_non_twinlift" not in _cached_summary:
        st.session_state.pop("hasil", None)
        st.warning(
            "⚠️ Hasil analisis sebelumnya sudah usang (aplikasi baru saja diperbarui). "
            "Silakan klik tombol di atas untuk menjalankan ulang analisis."
        )

if not run and "hasil" not in st.session_state:
    st.info("👆 Klik tombol di atas untuk menjalankan analisis pada file ini.")
    st.stop()

if run:
    try:
        with st.spinner("Memproses data..."):
            df = siapkan_data(raw, col_map, size_eligible)

            if len(df) == 0:
                st.error(
                    "❌ Setelah pembersihan, tidak ada baris data yang tersisa. "
                    "Kemungkinan kolom DISC_LOAD_TS & STACK_UNSTACK_TS yang dipilih "
                    "tidak berisi tanggal/jam yang valid, atau pemetaan kolom di "
                    "Langkah 2 salah. Silakan periksa kembali pemetaan kolom di atas."
                )
                st.stop()

            df_combo = layer1_combo(df, ambang_combo, size_eligible)
            twinlift_status_map, twinlift_gap_map = deteksi_twinlift(df_combo, ambang_twinlift, size_eligible)
            events = bentuk_event(df_combo)
            events["TWINLIFT_STATUS"] = events["GROUP_ID"].map(twinlift_status_map)
            events["TWINLIFT_GAP_MENIT"] = events["GROUP_ID"].map(twinlift_gap_map)
            events = layer2_dual(events, ambang_dual)
            events, event_id_map = beri_event_id(events, df_combo)
            out_df = gabungkan_hasil(df_combo, events, event_id_map)
            summary = hitung_ringkasan(events, out_df)
    except Exception as e:
        st.error(
            "❌ Terjadi error saat memproses data. Detail error ada di bawah ini — "
            "silakan screenshot/copy untuk dilaporkan."
        )
        with st.expander("Detail error (untuk dilaporkan)", expanded=True):
            st.exception(e)
        st.stop()

    st.session_state["hasil"] = {
        "out_df": out_df,
        "events": events,
        "summary": summary,
        "ambang_combo": ambang_combo,
        "ambang_dual": ambang_dual,
        "ambang_twinlift": ambang_twinlift,
    }
    # settingan yg dipakai run terakhir -- dipakai buat deteksi "pengaturan
    # sudah diubah tapi belum di-apply" (lihat blok peringatan di bawah)
    st.session_state["_ambang_terakhir"] = (ambang_combo, ambang_dual, ambang_twinlift)

hasil = st.session_state["hasil"]
out_df = hasil["out_df"]
events = hasil["events"]
summary = hasil["summary"]

# ----------------------------------------------------------------
# Peringatan kalau ambang batas di panel Pengaturan sudah diubah user
# TAPI tombol "Jalankan Analisis" belum diklik ulang -- supaya hasil
# yang ditampilkan tidak disangka pakai ambang yang baru padahal masih
# pakai ambang yang lama (sumber kebingungan paling umum).
# ----------------------------------------------------------------
_ambang_terakhir = st.session_state.get(
    "_ambang_terakhir", (hasil["ambang_combo"], hasil["ambang_dual"], hasil["ambang_twinlift"])
)
if (ambang_combo, ambang_dual, ambang_twinlift) != _ambang_terakhir:
    st.warning(
        "⚠️ Ambang batas di Pengaturan sudah diubah tapi belum diterapkan. "
        "Hasil di bawah masih pakai ambang yang lama — klik "
        "\"▶️ Jalankan Analisis Dual Cycle\" lagi untuk memperbarui."
    )

st.success("Analisis selesai!")

monthly = summary["monthly"].reset_index().rename(columns={"BULAN": "Bulan"})

# ================================================================
# DASHBOARD HASIL -- ditampilkan sbg 4 menu/tab: Dual Cycle,
# Twinlift, Analisis per Vessel, dan Download Hasil Analisis,
# langsung berisi grafik tanpa perlu scroll panjang lewat banyak
# subheader.
# ================================================================

tab_dual, tab_twinlift, tab_vessel, tab_download = st.tabs(
    ["📊 Dual Cycle", "🔗 Twinlift", "🚢 Per Vessel", "⬇️ Download Hasil Analisis"]
)

# ----------------------------------------------------------------
# TAB 1: DUAL CYCLE
# ----------------------------------------------------------------
with tab_dual:
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Event (Ritase)", summary["total_event"])
    k2.metric("Dual Cycle", summary["total_dual"])
    k3.metric("Non Dual", summary["total_single"])
    k4.metric("% Dual Cycle", f"{summary['pct_dual']*100:.1f}%")
    k5.metric("Jumlah Container - LOAD", summary["container_load"])
    k6.metric("Jumlah Container - DISC", summary["container_disc"])

    cc1, cc2 = st.columns(2)

    with cc1:
        pie_df = pd.DataFrame(
            {"Status": ["Dual Cycle", "Non Dual"], "Jumlah": [summary["total_dual"], summary["total_single"]]}
        )
        fig_pie = px.pie(
            pie_df, names="Status", values="Jumlah", hole=0.45,
            title="Dual Cycle vs Non Dual (berbasis Event)",
            color="Status",
            color_discrete_map={"Dual Cycle": "#2E86AB", "Non Dual": "#E76F51"},
        )
        fig_pie.update_traces(textinfo="percent+label")
        st.plotly_chart(fig_pie, width='stretch')

    with cc2:
        container_df = pd.DataFrame(
            {
                "Container": ["Combo", "Combo", "Single", "Single"],
                "Status": ["Dual Cycle", "Non Dual", "Dual Cycle", "Non Dual"],
                "Jumlah": [
                    summary["combo_dual"], summary["combo_single"],
                    summary["single_dual"], summary["single_single"],
                ],
            }
        )
        fig_bar = px.bar(
            container_df, x="Container", y="Jumlah", color="Status", barmode="group",
            title="Rincian Container x Status (Event)",
            color_discrete_map={"Dual Cycle": "#2E86AB", "Non Dual": "#E76F51"},
            text="Jumlah",
        )
        st.plotly_chart(fig_bar, width='stretch')

    if len(monthly) > 0:
        # Breakdown bulanan Dual Cycle vs Non Dual (%)
        monthly_dual_pct = monthly.melt(
            id_vars="Bulan",
            value_vars=["pct_dual", "pct_non_dual"],
            var_name="Kategori",
            value_name="Persentase",
        )
        monthly_dual_pct["Kategori"] = monthly_dual_pct["Kategori"].map(
            {"pct_dual": "Dual Cycle", "pct_non_dual": "Non Dual"}
        )
        monthly_dual_pct["Persentase"] = monthly_dual_pct["Persentase"] * 100

        fig_month_dual = px.bar(
            monthly_dual_pct, x="Bulan", y="Persentase", color="Kategori", barmode="stack",
            title="Breakdown Bulanan: Dual Cycle vs Non Dual (%)",
            color_discrete_map={"Dual Cycle": "#2E86AB", "Non Dual": "#E76F51"},
            text_auto=".1f",
        )
        fig_month_dual.update_layout(yaxis=dict(title="% dari Total Event", range=[0, 100]))
        st.plotly_chart(fig_month_dual, width='stretch')

        # Breakdown bulanan Combo vs Single (%)
        monthly_container_pct = monthly.melt(
            id_vars="Bulan",
            value_vars=["pct_combo", "pct_single"],
            var_name="Kategori",
            value_name="Persentase",
        )
        monthly_container_pct["Kategori"] = monthly_container_pct["Kategori"].map(
            {"pct_combo": "Combo", "pct_single": "Single"}
        )
        monthly_container_pct["Persentase"] = monthly_container_pct["Persentase"] * 100

        fig_month_container = px.bar(
            monthly_container_pct, x="Bulan", y="Persentase", color="Kategori", barmode="stack",
            title="Breakdown Bulanan: Combo vs Single (%)",
            color_discrete_map={"Combo": "#F4A261", "Single": "#8AB17D"},
            text_auto=".1f",
        )
        fig_month_container.update_layout(yaxis=dict(title="% dari Total Event", range=[0, 100]))
        st.plotly_chart(fig_month_container, width='stretch')

# ----------------------------------------------------------------
# TAB 2: TWINLIFT
# ----------------------------------------------------------------
with tab_twinlift:
    st.caption(
        "Twinlift = kontainer di dalam Combo 20ft, berasal dari kapal (VES_ID) & truk yang sama, "
        f"dengan jarak DISC_LOAD_TS antar keduanya ≤ {hasil['ambang_twinlift']:g} menit. "
        "Breakdown di bawah ini dihitung dari basis **Total Event**."
    )

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Total Event (Ritase)", summary["total_event"])
    t2.metric("Twinlift", summary["total_twinlift"])
    t3.metric("Bukan Twinlift", summary["total_non_twinlift"])
    t4.metric("% Twinlift dari Total Event", f"{summary['pct_twinlift_of_total']*100:.1f}%")

    tc1, tc2 = st.columns(2)

    with tc1:
        if summary["total_event"] > 0:
            twin_df = pd.DataFrame(
                {
                    "Status": ["Twinlift", "Bukan Twinlift"],
                    "Jumlah": [summary["total_twinlift"], summary["total_non_twinlift"]],
                }
            )
            fig_twin_pie = px.pie(
                twin_df, names="Status", values="Jumlah", hole=0.45,
                title="Twinlift vs Bukan Twinlift (dari Total Event)",
                color="Status",
                color_discrete_map={"Twinlift": "#6A4C93", "Bukan Twinlift": "#C0C0C0"},
            )
            fig_twin_pie.update_traces(textinfo="percent+label")
            st.plotly_chart(fig_twin_pie, width='stretch')
        else:
            st.info("Tidak ada event pada data ini.")

    with tc2:
        if len(monthly) > 0:
            monthly_twin_pct = monthly.melt(
                id_vars="Bulan",
                value_vars=["pct_twinlift", "pct_non_twinlift"],
                var_name="Kategori",
                value_name="Persentase",
            )
            monthly_twin_pct["Kategori"] = monthly_twin_pct["Kategori"].map(
                {"pct_twinlift": "Twinlift", "pct_non_twinlift": "Bukan Twinlift"}
            )
            monthly_twin_pct["Persentase"] = monthly_twin_pct["Persentase"] * 100

            fig_month_twin = px.bar(
                monthly_twin_pct, x="Bulan", y="Persentase", color="Kategori", barmode="stack",
                title="Breakdown Bulanan: Twinlift vs Bukan Twinlift (% dari Total Event)",
                color_discrete_map={"Twinlift": "#6A4C93", "Bukan Twinlift": "#C0C0C0"},
                text_auto=".1f",
            )
            fig_month_twin.update_layout(yaxis=dict(title="% dari Total Event", range=[0, 100]))
            st.plotly_chart(fig_month_twin, width='stretch')

# ----------------------------------------------------------------
# TAB 3: ANALISIS PER VESSEL
# ----------------------------------------------------------------
with tab_vessel:
    st.caption(
        "Cari satu kapal (VES_ID) untuk melihat karakteristiknya: Dual Cycle & Twinlift-nya "
        "dihitung dari seluruh aktivitas (baris LOAD/DISC) milik kapal tersebut."
    )

    # ------------------------------------------------------------
    # Jaring pengaman kedua: pastikan VES_ID benar-benar bersih
    # (dropna dulu, baru astype(str)+strip) sebelum di-sort. Kalau
    # langsung .astype(str) pada kolom yang masih mengandung NaN asli
    # (mis. dari file lama sebelum siapkan_data() diperbaiki, atau
    # dari sumber out_df lain), NaN bisa tetap jadi float, bukan
    # ter-konversi ke teks "nan" -- dan itu bikin sorted() gagal
    # dengan TypeError ('<' not supported between float dan str).
    # ------------------------------------------------------------
    ves_id_bersih = out_df["VES_ID"].dropna().astype(str).str.strip()
    ves_id_bersih = ves_id_bersih[ves_id_bersih != ""]
    vessel_options = sorted(ves_id_bersih.unique().tolist())

    if len(vessel_options) == 0:
        st.info("Tidak ada data VES_ID pada hasil analisis ini.")
    else:
        selected_vessel = st.selectbox("🔍 Cari / pilih VES_ID", vessel_options)

        vessel_df = out_df[out_df["VES_ID"].astype(str).str.strip() == selected_vessel]

        total_rec = len(vessel_df)
        dual_rec = int((vessel_df["STATUS"] == "Dual Cycle").sum())
        non_dual_rec = total_rec - dual_rec
        twinlift_rec = int((vessel_df["TWINLIFT_STATUS"] == "Twinlift").sum())
        non_twinlift_rec = total_rec - twinlift_rec
        combo_rec = int((vessel_df["CONTAINER_STATUS"] == "Combo").sum())
        single_rec = total_rec - combo_rec

        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Total Aktivitas (baris)", total_rec)
        v2.metric("Dual Cycle", dual_rec)
        v3.metric("% Dual Cycle", f"{(dual_rec/total_rec*100) if total_rec else 0:.1f}%")
        v4.metric("Twinlift", twinlift_rec)
        v5.metric("% Twinlift", f"{(twinlift_rec/total_rec*100) if total_rec else 0:.1f}%")

        vc1, vc2 = st.columns(2)
        with vc1:
            if total_rec > 0:
                pie_dual_v = pd.DataFrame(
                    {"Status": ["Dual Cycle", "Non Dual"], "Jumlah": [dual_rec, non_dual_rec]}
                )
                fig_v1 = px.pie(
                    pie_dual_v, names="Status", values="Jumlah", hole=0.45,
                    title=f"Dual Cycle vs Non Dual — {selected_vessel}",
                    color="Status",
                    color_discrete_map={"Dual Cycle": "#2E86AB", "Non Dual": "#E76F51"},
                )
                fig_v1.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_v1, width='stretch')

        with vc2:
            if total_rec > 0:
                pie_twin_v = pd.DataFrame(
                    {"Status": ["Twinlift", "Bukan Twinlift"], "Jumlah": [twinlift_rec, non_twinlift_rec]}
                )
                fig_v2 = px.pie(
                    pie_twin_v, names="Status", values="Jumlah", hole=0.45,
                    title=f"Twinlift vs Bukan Twinlift — {selected_vessel}",
                    color="Status",
                    color_discrete_map={"Twinlift": "#6A4C93", "Bukan Twinlift": "#C0C0C0"},
                )
                fig_v2.update_traces(textinfo="percent+label")
                st.plotly_chart(fig_v2, width='stretch')

        vcont_df = pd.DataFrame(
            {"Container": ["Combo", "Single"], "Jumlah": [combo_rec, single_rec]}
        )
        fig_v3 = px.bar(
            vcont_df, x="Container", y="Jumlah",
            title=f"Combo vs Single — {selected_vessel}",
            color="Container",
            color_discrete_map={"Combo": "#F4A261", "Single": "#8AB17D"},
            text="Jumlah",
        )
        st.plotly_chart(fig_v3, width='stretch')

# ----------------------------------------------------------------
# TAB 4: DOWNLOAD HASIL ANALISIS
# ----------------------------------------------------------------
with tab_download:
    st.write(f"Data hasil analisis ({len(out_df)} baris):")
    # Preview dibatasi (bukan full 500rb baris) supaya rendering tabel ini
    # tidak ikut membebani SETIAP rerun app -- lihat catatan di bawah soal
    # kenapa semua isi tab dieksekusi ulang di tiap interaksi. Data lengkap
    # tetap bisa didapat lewat tombol download di bawah.
    st.dataframe(out_df.head(1000), width='stretch', height=400)
    if len(out_df) > 1000:
        st.caption(f"Menampilkan 1.000 baris pertama dari {len(out_df)} baris. Download file di bawah untuk data lengkap.")

    # ------------------------------------------------------------
    # PENTING -- pembuatan file (CSV/Excel) dari data besar (bisa
    # ratusan ribu baris) TIDAK BOLEH dijalankan otomatis di setiap
    # rerun script. Streamlit menjalankan ULANG SELURUH ISI SEMUA TAB
    # (bukan cuma tab yang lagi dibuka) setiap kali ada interaksi apa
    # pun di app -- termasuk ganti pilihan VES_ID di tab "Per Vessel".
    # Kalau data di-generate ulang langsung di sini tanpa penjagaan,
    # SETIAP interaksi (ganti dropdown, dll) akan memicu penulisan
    # ulang file, dan itu bisa bikin app lambat / boros memori.
    #
    # Solusinya: CSV & Excel dibangun cuma SEKALI per hasil analisis
    # (di-cache di session_state, dikunci oleh signature dari hasil),
    # bukan dibangun ulang di setiap rerun selama hasilnya belum ganti.
    # Excel di sini SENGAJA versi ringan (data saja, tanpa sheet
    # Ringkasan & tanpa chart) supaya proses download tetap cepat.
    # ------------------------------------------------------------
    hasil_sig = (
        hasil["ambang_combo"], hasil["ambang_dual"], hasil["ambang_twinlift"], len(out_df),
    )
    if st.session_state.get("_download_sig") != hasil_sig:
        st.session_state.pop("_csv_bytes", None)
        st.session_state.pop("_excel_bytes", None)
        st.session_state["_download_sig"] = hasil_sig

    if "_csv_bytes" not in st.session_state:
        st.session_state["_csv_bytes"] = out_df.to_csv(index=False).encode("utf-8-sig")
    if "_excel_bytes" not in st.session_state:
        with st.spinner("Menyiapkan file Excel (data saja, ringan & cepat)..."):
            st.session_state["_excel_bytes"] = build_excel_data_only(out_df)

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            "⬇️ Download CSV (Data)",
            data=st.session_state["_csv_bytes"],
            file_name="Hasil_Analisis_Dual_Cycle.csv",
            mime="text/csv",
            width='stretch',
        )
    with dcol2:
        st.download_button(
            "⬇️ Download Excel (Data)",
            data=st.session_state["_excel_bytes"],
            file_name="Hasil_Analisis_Dual_Cycle.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )
