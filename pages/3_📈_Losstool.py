# -*- coding: utf-8 -*-
"""
loss_processor.py
==================
Cable Loss XML 轉換核心邏輯（與 Streamlit UI 完全分離）。

- 使用 xml.etree.ElementTree 解析/修改 XML，取代原本容易失效的正規表達式作法。
- 提供 Excel 資料驗證 (validate_dataframe)。
- 支援多通道（Main1/Main2/AUX...）批次比對與補償。

這個模組不依賴 streamlit，因此可以獨立被 CLI 工具或單元測試呼叫。
"""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

# ============================================================
# 1. 內建預設 XML 模板
# ============================================================

DEFAULT_WWAN_XML = """<CableLossConfig>
  <Instr Instrument="MT8821C">
    <Module>MT8821C</Module>
  </Instr>
  <Cfg Option="External">
    <Std Val="LTE">
      <Data>
        <Main1_DL>0</Main1_DL>
        <Main1_UL>0</Main1_UL>
        <AUX1_DL>0</AUX1_DL>
        <SingleBoxSlot2_TRx1MainDL>0</SingleBoxSlot2_TRx1MainDL>
        <SingleBoxSlot2_TRx1MainUL>0</SingleBoxSlot2_TRx1MainUL>
        <SingleBoxSlot2_TRx1AuxDL>0</SingleBoxSlot2_TRx1AuxDL>
        <SingleBoxSlot2_TRx2MainDL>0</SingleBoxSlot2_TRx2MainDL>
        <SingleBoxSlot2_TRx2MainUL>0</SingleBoxSlot2_TRx2MainUL>
        <SingleBoxSlot2_TRx2AuxDL>0</SingleBoxSlot2_TRx2AuxDL>
        <SingleBoxSlot2_TRx3MainDL>0</SingleBoxSlot2_TRx3MainDL>
        <SingleBoxSlot2_TRx3MainUL>0</SingleBoxSlot2_TRx3MainUL>
        <SingleBoxSlot2_TRx3AuxDL>0</SingleBoxSlot2_TRx3AuxDL>
        <SingleBoxSlot2_TRx4MainDL>0</SingleBoxSlot2_TRx4MainDL>
        <SingleBoxSlot2_TRx4MainUL>0</SingleBoxSlot2_TRx4MainUL>
        <SingleBoxSlot2_TRx4AuxDL>0</SingleBoxSlot2_TRx4AuxDL>
        <EntryModeSlot1_TRx2MainDL>0</EntryModeSlot1_TRx2MainDL>
        <EntryModeSlot1_TRx2MainUL>0</EntryModeSlot1_TRx2MainUL>
        <EntryModeSlot1_TRx2AuxDL>0</EntryModeSlot1_TRx2AuxDL>
      </Data>
    </Std>
  </Cfg>
  <Cfg Option="Common">
    <Std Val="LTE">
      <Data>
        <Frequency>622</Frequency>
        <Main1_DL>47.5</Main1_DL>
        <Main1_UL>40.0</Main1_UL>
        <Main2_DL>40</Main2_DL>
        <Main2_UL>40</Main2_UL>
        <AUX1_DL>40</AUX1_DL>
        <AUX2_DL>40</AUX2_DL>
        <AUX3_DL>40</AUX3_DL>
        <AUX4_DL>40</AUX4_DL>
      </Data>
      <Data>
        <Frequency>634.5</Frequency>
        <Main1_DL>49.5</Main1_DL>
        <Main1_UL>40.0</Main1_UL>
        <Main2_DL>40</Main2_DL>
        <Main2_UL>40</Main2_UL>
        <AUX1_DL>40</AUX1_DL>
        <AUX2_DL>40</AUX2_DL>
        <AUX3_DL>40</AUX3_DL>
        <AUX4_DL>40</AUX4_DL>
      </Data>
    </Std>
    <Std Val="LTE_P2" />
    <Std Val="LTE_SingleBox" />
    <Std Val="LTE_EntryMode" />
  </Cfg>
</CableLossConfig>"""

DEFAULT_WLAN_XML = """<CableLossConfig>
  <Instrument>MT8862A SiSo</Instrument>
  <Config Option="Master">
    <Table Val="1">
      <Data>
        <Frequency>2412</Frequency>
        <Main1_IN>49</Main1_IN>
        <Main1_OUT>49</Main1_OUT>
        <Main2_IN>0</Main2_IN>
        <Main2_OUT>0</Main2_OUT>
        <AUX_OUT>0</AUX_OUT>
      </Data>
      <Data>
        <Frequency>2437</Frequency>
        <Main1_IN>49</Main1_IN>
        <Main1_OUT>49</Main1_OUT>
        <Main2_IN>0</Main2_IN>
        <Main2_OUT>0</Main2_OUT>
        <AUX_OUT>0</AUX_OUT>
      </Data>
    </Table>
  </Config>
</CableLossConfig>"""

# ============================================================
# 2. 通道 (Channel) 定義 -- 支援多通道批次比對
# ============================================================
# 每個 mode 定義：
#   primary : 使用者「一定要」在 Excel 提供的欄位（沒有就視為錯誤/0）
#   extra   : 選填欄位，若 Excel 有提供該欄則覆蓋，否則沿用 XML 原值或預設值
#   defaults: extra 欄位在「新增全新頻點」時的預設值

CHANNEL_CONFIG = {
    "WWAN": {
        "primary": ["Main1_DL", "Main1_UL"],
        "extra": ["Main2_DL", "Main2_UL", "AUX1_DL", "AUX2_DL", "AUX3_DL", "AUX4_DL"],
        "extra_default": "40",
        "primary_default": "0",
    },
    "WLAN": {
        "primary": ["Main1_IN", "Main1_OUT"],
        "extra": ["Main2_IN", "Main2_OUT", "AUX_OUT"],
        "extra_default": "0",
        "primary_default": "0",
    },
}


def all_channels(mode_key: str) -> list[str]:
    """回傳該 mode 完整、有順序的通道欄位清單（Frequency 除外）。"""
    cfg = CHANNEL_CONFIG[mode_key]
    return list(cfg["primary"]) + list(cfg["extra"])


# ============================================================
# 2b. 通道角色 (TX/RX) -- 決定自動補償時 Diff 的計算方向
# ============================================================
# tx (上行/發射，例如 Main1_UL)：Diff = Target - Measured   （對應 TRP DIFF）
# rx (下行/接收，例如 Main1_DL)：Diff = Measured - Target   （對應 TIS DIFF）

CHANNEL_ROLES = {
    "WWAN": {"Main1_DL": "rx", "Main1_UL": "tx"},
    "WLAN": {"Main1_IN": "rx", "Main1_OUT": "tx"},
}

# 目標值的預設值（可在 UI 上手動調整/覆蓋，這裡只是建表時的初始值）
DEFAULT_ROLE_TARGET = {"tx": 23.0, "rx": -101.0}  # TRP 預設 23 dB / TIS 預設 -101 dBm


def channel_role_label(mode_key: str, channel: str) -> str:
    role = CHANNEL_ROLES.get(mode_key, {}).get(channel, "rx")
    return "TX / 發射 (TRP)" if role == "tx" else "RX / 接收 (TIS)"


# ============================================================
# 3. 頻率正規化
# ============================================================

def normalize_freq(val) -> str:
    """把 622 / 622.0 / '622.00' 等都正規化成同一個字串 key，方便比對。"""
    try:
        f = float(val)
        s = f"{f:.6f}".rstrip("0").rstrip(".")
        return s if s else "0"
    except (TypeError, ValueError):
        return str(val).strip()


# ============================================================
# 4. Excel 資料驗證
# ============================================================

@dataclass
class ValidationResult:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return len(self.errors) == 0


def validate_dataframe(df: pd.DataFrame, mode_key: str) -> ValidationResult:
    """檢查必要欄位是否存在、數值是否合法，回傳可直接顯示給使用者的訊息。"""
    result = ValidationResult()
    cfg = CHANNEL_CONFIG[mode_key]
    required_cols = ["Frequency"] + cfg["primary"]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        result.errors.append(
            f"Excel 缺少必要欄位：{', '.join(missing)}（{mode_key} 模式需要 {', '.join(required_cols)}）"
        )
        return result  # 欄位都不齊，後面逐列檢查沒有意義

    if len(df) == 0:
        result.warnings.append("Excel 目前沒有任何資料列。")
        return result

    seen_freqs = {}
    for idx, row in df.iterrows():
        line_no = idx + 2  # Excel 從第 1 列是標題，資料從第 2 列開始
        freq_raw = row.get("Frequency")

        if pd.isna(freq_raw) or str(freq_raw).strip() == "":
            continue  # 空白列交給呼叫端跳過，不視為錯誤

        try:
            freq_val = float(freq_raw)
        except (TypeError, ValueError):
            result.errors.append(f"第 {line_no} 列：Frequency「{freq_raw}」不是合法數字")
            continue

        norm = normalize_freq(freq_val)
        if norm in seen_freqs:
            result.warnings.append(
                f"第 {line_no} 列與第 {seen_freqs[norm]} 列的 Frequency 重複（{norm}），將以較後面的列為準"
            )
        seen_freqs[norm] = line_no

        for col in cfg["primary"]:
            val = row.get(col)
            if pd.isna(val) or str(val).strip() == "":
                result.warnings.append(f"第 {line_no} 列：{col} 未填寫，將以 0 帶入")
                continue
            try:
                float(val)
            except (TypeError, ValueError):
                result.errors.append(f"第 {line_no} 列：{col}「{val}」不是合法數字")

        for col in cfg["extra"]:
            if col in df.columns:
                val = row.get(col)
                if pd.notna(val) and str(val).strip() != "":
                    try:
                        float(val)
                    except (TypeError, ValueError):
                        result.errors.append(f"第 {line_no} 列：{col}「{val}」不是合法數字")

    return result


# ============================================================
# 4b. 功能 1：統一 Loss 數值 (批次帶入同一個值，可調)
# ============================================================

def apply_uniform_loss(df: pd.DataFrame, columns: list[str], value: float) -> pd.DataFrame:
    """把指定的通道欄位全部覆蓋成同一個數值，回傳新的 DataFrame（不改動原本的 df）。"""
    result = df.copy()
    for c in columns:
        if c in result.columns:
            result[c] = float(value)  # 直接整欄改型別成 float 並覆蓋成同一個值
    return result


# ============================================================
# 4c. 功能 2：自動補償 (原Loss / 測量值 / 目標值 / 差值 / 補償後 Loss)
# ============================================================
# 邏輯與 WWAN_Lossdata.xlsx 的 Data 分頁一致：
#   TX (上行/TRP)：Diff = Target - Measured
#   RX (下行/TIS)：Diff = Measured - Target
#   補償後 Loss = 原 Loss + Diff

def compute_diff(role: str, target: float, measured: float) -> float:
    if role == "tx":
        return target - measured
    return measured - target


def compute_compensated_loss(original_loss: float, diff: float) -> float:
    return original_loss + diff


def build_compensation_frame(
    working_df: pd.DataFrame, mode_key: str, default_targets: "dict[str, float] | None" = None
) -> pd.DataFrame:
    """依照目前的 Loss 表格（Frequency + 主要通道）建立一份全新的補償試算表。

    default_targets：可選，依角色（"tx"/"rx"）指定目標值的初始值。
    沒有指定的話，TX（TRP）預設 23、RX（TIS）預設 -101，之後仍可在表格內手動修改。
    """
    targets = dict(DEFAULT_ROLE_TARGET)
    if default_targets:
        targets.update(default_targets)

    cfg = CHANNEL_CONFIG[mode_key]
    comp = pd.DataFrame()
    comp["Frequency"] = working_df.get("Frequency", pd.Series(dtype=object))
    for ch in cfg["primary"]:
        role = CHANNEL_ROLES.get(mode_key, {}).get(ch, "rx")
        comp[f"{ch}_OriginalLoss"] = pd.to_numeric(working_df.get(ch), errors="coerce") if ch in working_df.columns else None
        comp[f"{ch}_Target"] = targets.get(role, pd.NA)
        comp[f"{ch}_Measured"] = pd.NA
        comp[f"{ch}_Diff"] = pd.NA
        comp[f"{ch}_CompLoss"] = pd.NA
    return comp


def recompute_compensation(comp_df: pd.DataFrame, mode_key: str) -> pd.DataFrame:
    """依照目前填好的 Target / Measured / OriginalLoss，重新計算 Diff 與補償後 Loss。"""
    cfg = CHANNEL_CONFIG[mode_key]
    comp_df = comp_df.copy()
    for ch in cfg["primary"]:
        role = CHANNEL_ROLES.get(mode_key, {}).get(ch, "rx")
        target_col, measured_col = f"{ch}_Target", f"{ch}_Measured"
        orig_col, diff_col, comp_col = f"{ch}_OriginalLoss", f"{ch}_Diff", f"{ch}_CompLoss"
        for col in (target_col, measured_col, orig_col, diff_col, comp_col):
            if col not in comp_df.columns:
                comp_df[col] = pd.NA

        for idx in comp_df.index:
            t = comp_df.at[idx, target_col]
            m = comp_df.at[idx, measured_col]
            o = comp_df.at[idx, orig_col]
            if pd.notna(t) and pd.notna(m):
                diff = compute_diff(role, float(t), float(m))
                comp_df.at[idx, diff_col] = diff
                if pd.notna(o):
                    comp_df.at[idx, comp_col] = compute_compensated_loss(float(o), diff)
                else:
                    comp_df.at[idx, comp_col] = pd.NA
            else:
                comp_df.at[idx, diff_col] = pd.NA
                comp_df.at[idx, comp_col] = pd.NA
    return comp_df


def merge_measured_file(comp_df: pd.DataFrame, measured_df: pd.DataFrame, mode_key: str) -> pd.DataFrame:
    """把外部量測數據檔案（需含 Frequency + 通道欄位）合併進補償試算表的 Measured 欄位。"""
    if "Frequency" not in measured_df.columns:
        raise ValueError("量測數據檔案缺少必要欄位：Frequency")

    cfg = CHANNEL_CONFIG[mode_key]
    comp_df = comp_df.copy()
    measured_df = measured_df.copy()
    measured_df["_norm_freq"] = measured_df["Frequency"].apply(normalize_freq)
    lookup = measured_df.drop_duplicates("_norm_freq", keep="last").set_index("_norm_freq")

    matched = 0
    for ch in cfg["primary"]:
        if ch not in measured_df.columns:
            continue
        measured_col = f"{ch}_Measured"
        if measured_col not in comp_df.columns:
            comp_df[measured_col] = pd.NA
        for idx, row in comp_df.iterrows():
            nf = normalize_freq(row["Frequency"])
            if nf in lookup.index:
                val = lookup.loc[nf, ch]
                if pd.notna(val):
                    comp_df.at[idx, measured_col] = val
                    matched += 1

    return comp_df, matched


def apply_compensation_to_working(working_df: pd.DataFrame, comp_df: pd.DataFrame, mode_key: str) -> tuple[pd.DataFrame, int]:
    """把補償試算表算出來的「補償後 Loss」寫回主要的 Loss 表格（依 Frequency 比對）。"""
    cfg = CHANNEL_CONFIG[mode_key]
    working_df = working_df.copy()
    for ch in cfg["primary"]:
        if ch in working_df.columns:
            working_df[ch] = pd.to_numeric(working_df[ch], errors="coerce").astype("float64")  # 強制轉成 float，避免整數欄位塞不進小數

    comp_lookup = {normalize_freq(f): row for f, row in zip(comp_df["Frequency"], comp_df.to_dict("records"))}

    applied = 0
    for idx, row in working_df.iterrows():
        nf = normalize_freq(row.get("Frequency"))
        comp_row = comp_lookup.get(nf)
        if not comp_row:
            continue
        for ch in cfg["primary"]:
            comp_val = comp_row.get(f"{ch}_CompLoss")
            if pd.notna(comp_val):
                working_df.at[idx, ch] = round(float(comp_val), 3)
                applied += 1

    return working_df, applied


# ============================================================
# 4c-2. 功能 2b：以量產測試結果 (MVO5 CSV) + Frequency List 自動比對補償 (目前僅支援 WWAN)
# ============================================================
# 流程：
#   1. 讀取 MVO5 測試結果 CSV：每一列有 Band / Channel / Result(TIS 測試值) /
#      Description（其中內嵌 "Tx power = xx.x" 為 TRP 測試值）。
#   2. 讀取 Frequency List Excel：每一列有 Band / Channel / Frequency (MHz) / TX/RX(角色)，
#      用 (Band, Channel) 查出該測試點對應的實際 Frequency 與角色 (RX / TX / TXRX)。
#   3. 依角色決定要用 TIS 或 TRP 的哪個測試值來計算補償：
#        角色 RX   -> 用 Result(TIS) 對比 RX 目標值，補償結果只寫回 Main1_DL（Main1_UL 不變）
#        角色 TX   -> 用 Tx power(TRP) 對比 TX 目標值，補償結果只寫回 Main1_UL（Main1_DL 不變）
#        角色 TXRX -> 兩者都計算，Main1_DL / Main1_UL 都更新
#   4. 若多個 Band 換算出同一個 Frequency（頻率重疊），預設只勾選 Band 數字較小的那筆，
#      使用者仍可在預覽表格中手動勾選/取消。

MVO5_TXPOWER_RE = re.compile(r"Tx\s*power\s*=\s*([-+]?\d+(?:\.\d+)?)", re.IGNORECASE)


def parse_mvo5_result_csv(df: pd.DataFrame) -> pd.DataFrame:
    """解析 MVO5 test-for-loss 測試結果 CSV。

    需要欄位：Band, Channel, Result, Description。
    回傳欄位：Band, Channel, TIS（Result 數值), TRP（從 Description 解析出的 Tx power 數值）。
    """
    required = ["Band", "Channel", "Result", "Description"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"測試結果 CSV 缺少必要欄位：{', '.join(missing)}")

    out = pd.DataFrame()
    out["Band"] = df["Band"].astype(str).str.strip()
    out["Channel"] = pd.to_numeric(df["Channel"], errors="coerce")
    out["TIS"] = pd.to_numeric(df["Result"], errors="coerce")

    def extract_txpower(desc):
        if pd.isna(desc):
            return None
        m = MVO5_TXPOWER_RE.search(str(desc))
        return float(m.group(1)) if m else None

    out["TRP"] = df["Description"].apply(extract_txpower)
    out = out.dropna(subset=["Channel"]).copy()
    out["Channel"] = out["Channel"].astype("int64")
    return out.reset_index(drop=True)


def load_frequency_list(freq_df: pd.DataFrame) -> pd.DataFrame:
    """整理 Frequency List Excel，標準化欄位名稱為 Band / Channel / Frequency / Role。

    來源欄位預期為：Band, Channel, Frequency (MHz), TX/RX（欄位名稱大小寫/空白容忍度較高）。
    Role 會標準化成大寫的 "RX" / "TX" / "TXRX"。
    """
    colmap = {}
    for c in freq_df.columns:
        cs = str(c).strip()
        if cs.lower().startswith("frequency"):
            colmap[c] = "Frequency"
        elif cs.strip().upper().replace(" ", "") == "TX/RX":
            colmap[c] = "Role"
        else:
            colmap[c] = cs
    out = freq_df.rename(columns=colmap).copy()

    required = ["Band", "Channel", "Frequency", "Role"]
    missing = [c for c in required if c not in out.columns]
    if missing:
        raise ValueError(f"Frequency List 缺少必要欄位：{', '.join(missing)}")

    out["Band"] = out["Band"].astype(str).str.strip()
    out["Channel"] = pd.to_numeric(out["Channel"], errors="coerce")
    out = out.dropna(subset=["Channel"]).copy()
    out["Channel"] = out["Channel"].astype("int64")
    out["Frequency"] = pd.to_numeric(out["Frequency"], errors="coerce")
    out["Role"] = out["Role"].astype(str).str.strip().str.upper()
    return out[["Band", "Channel", "Frequency", "Role"]].reset_index(drop=True)


def _band_number(band: str) -> float:
    """從 'B71' 這種字串取出數字部分，方便比較 Band 大小；解析失敗回傳 inf（視為最大，排到最後）。"""
    m = re.search(r"(\d+)", str(band))
    return float(m.group(1)) if m else float("inf")


def filter_by_band_list(df: pd.DataFrame, band_list: "list[str] | None", band_col: str = "Band") -> pd.DataFrame:
    """依 Band 清單篩選 DataFrame（不分大小寫比對 "B1"/"b1" 皆可）。

    band_list 為 None 或空清單時，直接回傳原始 df（不篩選，維持舊行為）。
    這個函式用在：當需要補償的 Band 數量較少時（例如 Bandlist.txt 只列出 B1、B2），
    先用 Band 清單縮小 Frequency List／MVO5 測試結果的範圍，只針對這些 Band 對應的
    Channel/Frequency 做比對與補償，避免掃描不相關的 Band 資料。
    """
    if not band_list or band_col not in df.columns:
        return df
    normalized = {str(b).strip().upper() for b in band_list}
    return df[df[band_col].astype(str).str.strip().str.upper().isin(normalized)].copy()


def build_mvo5_compensation_table(
    mvo5_df: pd.DataFrame,
    freq_list_df: pd.DataFrame,
    working_df: pd.DataFrame,
    targets: "dict[str, float] | None" = None,
    band_list: "list[str] | None" = None,
) -> pd.DataFrame:
    """把 MVO5 測試結果依 Channel 查 Frequency List，取得 Frequency / Role，
    再依角色計算 Diff / 補償後 Loss，回傳可供預覽（含 Band 欄位）與勾選的表格。

    band_list：可選。若提供（例如從 Bandlist.txt 解析出的 ["B1", "B2"]），代表這次只需要
      補償這幾個 Band，會先用它分別篩選 freq_list_df 與 mvo5_df，只保留這些 Band 的資料再比對，
      而不是掃描 Frequency List / MVO5 CSV 裡的全部 Band。當 Band 數量遠少於 Frequency List
      涵蓋的 Band 數時，可大幅縮小比對範圍、加快處理速度，也讓輸出的補償表只包含關心的 Band。
      不提供（None）時維持原本行為：使用 MVO5 CSV 裡出現的所有 Band。

    輸出欄位：Band, Channel, Frequency, Role,
      Main1_DL_OriginalLoss/Target/Measured/Diff/CompLoss,
      Main1_UL_OriginalLoss/Target/Measured/Diff/CompLoss,
      Overlap（是否與其他 Band 共用同一個 Frequency）, Selected（預設勾選狀態）。
    """
    targets_ = dict(DEFAULT_ROLE_TARGET)
    if targets:
        targets_.update(targets)

    # 若指定了 Band 清單，先縮小 Frequency List 與 MVO5 測試結果的範圍到這幾個 Band。
    freq_list_df = filter_by_band_list(freq_list_df, band_list, band_col="Band")
    mvo5_df = filter_by_band_list(mvo5_df, band_list, band_col="Band")

    if band_list and freq_list_df.empty:
        raise ValueError(
            f"Band List {band_list} 在 Frequency List 中找不到對應的 Band，請確認拼字或 Frequency List 內容。"
        )
    if band_list and mvo5_df.empty:
        raise ValueError(
            f"Band List {band_list} 在 MVO5 測試結果 CSV 中找不到對應的 Band，請確認測試結果檔案內容。"
        )

    # Channel 在 Frequency List 中應為唯一值，直接用 Channel 查表（Band 僅作為顯示/比對用）。
    freq_lookup = freq_list_df.drop_duplicates("Channel", keep="last").set_index("Channel")

    work = working_df.copy()
    work["_norm_freq"] = work["Frequency"].apply(normalize_freq)
    orig_lookup = work.drop_duplicates("_norm_freq", keep="last").set_index("_norm_freq")

    def orig_loss(norm_freq, ch):
        if norm_freq not in orig_lookup.index:
            return None
        v = orig_lookup.loc[norm_freq].get(ch)
        return float(v) if pd.notna(v) else None

    rows = []
    for _, r in mvo5_df.iterrows():
        channel = int(r["Channel"])
        if channel not in freq_lookup.index:
            continue
        fl = freq_lookup.loc[channel]
        freq_val = float(fl["Frequency"])
        role = str(fl["Role"]).upper()
        norm_freq = normalize_freq(freq_val)

        row = {
            "Band": r["Band"],
            "Channel": channel,
            "Frequency": freq_val,
            "Role": role,
        }

        dl_ol = orig_loss(norm_freq, "Main1_DL")
        if role in ("RX", "TXRX") and pd.notna(r.get("TIS")):
            measured = float(r["TIS"])
            diff = compute_diff("rx", targets_["rx"], measured)
            comp = compute_compensated_loss(dl_ol, diff) if dl_ol is not None else None
            row.update({
                "Main1_DL_OriginalLoss": dl_ol, "Main1_DL_Target": targets_["rx"],
                "Main1_DL_Measured": measured, "Main1_DL_Diff": diff, "Main1_DL_CompLoss": comp,
            })
        else:
            row.update({
                "Main1_DL_OriginalLoss": dl_ol, "Main1_DL_Target": None,
                "Main1_DL_Measured": None, "Main1_DL_Diff": None, "Main1_DL_CompLoss": None,
            })

        ul_ol = orig_loss(norm_freq, "Main1_UL")
        if role in ("TX", "TXRX") and pd.notna(r.get("TRP")):
            measured = float(r["TRP"])
            diff = compute_diff("tx", targets_["tx"], measured)
            comp = compute_compensated_loss(ul_ol, diff) if ul_ol is not None else None
            row.update({
                "Main1_UL_OriginalLoss": ul_ol, "Main1_UL_Target": targets_["tx"],
                "Main1_UL_Measured": measured, "Main1_UL_Diff": diff, "Main1_UL_CompLoss": comp,
            })
        else:
            row.update({
                "Main1_UL_OriginalLoss": ul_ol, "Main1_UL_Target": None,
                "Main1_UL_Measured": None, "Main1_UL_Diff": None, "Main1_UL_CompLoss": None,
            })

        rows.append(row)

    result = pd.DataFrame(rows)
    if result.empty:
        for c in ("Overlap", "Selected"):
            result[c] = []
        return result

    # 頻率重疊處理：同一個 Frequency 若對應到多個不同 Band，預設只勾選 Band 數字最小的那筆。
    result["_freq_norm"] = result["Frequency"].apply(normalize_freq)
    result["_band_num"] = result["Band"].apply(_band_number)

    winner_band_num = result.groupby("_freq_norm")["_band_num"].min()
    band_counts = result.groupby("_freq_norm")["Band"].nunique()

    result["Overlap"] = result["_freq_norm"].map(lambda f: band_counts.get(f, 1) > 1)
    result["Selected"] = result.apply(
        lambda row: row["_band_num"] <= winner_band_num.get(row["_freq_norm"], row["_band_num"]),
        axis=1,
    )

    result = result.drop(columns=["_freq_norm", "_band_num"])
    result = result.sort_values(["Frequency", "Band"]).reset_index(drop=True)
    return result


def apply_mvo5_compensation_to_working(
    working_df: pd.DataFrame, mvo5_comp_df: pd.DataFrame
) -> "tuple[pd.DataFrame, int, list[str]]":
    """把 MVO5 補償試算表中「已勾選 (Selected)」的列，依 Frequency 寫回主要的 Loss 表格。

    Role=RX 只更新 Main1_DL（Main1_UL 維持原數值）；
    Role=TX 只更新 Main1_UL（Main1_DL 維持原數值）；
    Role=TXRX 則 Main1_DL / Main1_UL 都更新。

    只會更新「目前 Loss 表格中已經存在」的頻率點，不會因為測試結果多出新的頻率點就自動新增列，
    確保最終輸出的 XML 仍然只包含使用者原本輸入（Excel / XML / Band List）的頻率點。
    若某個勾選的頻率點在目前 Loss 表格中找不到，會被跳過，並回傳在 skipped 清單中方便提醒使用者。
    """
    working_df = working_df.copy()
    for ch in ("Main1_DL", "Main1_UL"):
        if ch in working_df.columns:
            working_df[ch] = pd.to_numeric(working_df[ch], errors="coerce").astype("float64")

    if "Selected" not in mvo5_comp_df.columns or len(mvo5_comp_df) == 0:
        return working_df, 0, []

    selected = mvo5_comp_df[mvo5_comp_df["Selected"] == True]  # noqa: E712
    applied = 0
    skipped: list[str] = []

    for _, row in selected.iterrows():
        norm_freq = normalize_freq(row["Frequency"])
        mask = working_df["Frequency"].apply(normalize_freq) == norm_freq
        if not mask.any():
            skipped.append(
                f"Band {row.get('Band')} / Channel {row.get('Channel')} / {row.get('Frequency')} MHz"
                "：目前 Loss 表格中沒有這個頻率點，已略過（不會自動新增）。"
            )
            continue

        role = row["Role"]
        if role in ("RX", "TXRX") and pd.notna(row.get("Main1_DL_CompLoss")):
            working_df.loc[mask, "Main1_DL"] = round(float(row["Main1_DL_CompLoss"]), 3)
            applied += 1
        if role in ("TX", "TXRX") and pd.notna(row.get("Main1_UL_CompLoss")):
            working_df.loc[mask, "Main1_UL"] = round(float(row["Main1_UL_CompLoss"]), 3)
            applied += 1

    return working_df, applied, skipped


def mvo5_compensation_warnings(mvo5_comp_df: pd.DataFrame) -> list[str]:
    """檢查 MVO5 補償試算表，列出「缺原Loss」或「缺測量值/無法算出補償」的列，方便在畫面上提醒使用者。

    - 缺原Loss：該角色需要的 Main1_DL/Main1_UL 原Loss 是 None（代表這個 Frequency 在目前的 Loss 表格中不存在）。
    - 缺測量值/補償：該角色需要用到的測量值是 None，或即使有原Loss/測量值但算不出 CompLoss。
    """
    warnings: list[str] = []
    if mvo5_comp_df is None or len(mvo5_comp_df) == 0:
        return warnings

    for _, row in mvo5_comp_df.iterrows():
        role = str(row.get("Role", "")).upper()
        label = f"Band {row.get('Band')} / Channel {row.get('Channel')} / {row.get('Frequency')} MHz"

        if role in ("RX", "TXRX"):
            if pd.isna(row.get("Main1_DL_OriginalLoss")):
                warnings.append(f"⚠️ {label}：Loss 表格中找不到這個頻率，缺少 DL 原Loss，無法計算補償。")
            elif pd.isna(row.get("Main1_DL_Measured")):
                warnings.append(f"⚠️ {label}：缺少 DL (TIS) 測量值，無法計算補償。")
            elif pd.isna(row.get("Main1_DL_CompLoss")):
                warnings.append(f"⚠️ {label}：DL 補償後 Loss 無法計算（請檢查原Loss/測量值/目標值）。")

        if role in ("TX", "TXRX"):
            if pd.isna(row.get("Main1_UL_OriginalLoss")):
                warnings.append(f"⚠️ {label}：Loss 表格中找不到這個頻率，缺少 UL 原Loss，無法計算補償。")
            elif pd.isna(row.get("Main1_UL_Measured")):
                warnings.append(f"⚠️ {label}：缺少 UL (TRP) 測量值，無法計算補償。")
            elif pd.isna(row.get("Main1_UL_CompLoss")):
                warnings.append(f"⚠️ {label}：UL 補償後 Loss 無法計算（請檢查原Loss/測量值/目標值）。")

    return warnings


# ============================================================
# 4d. 功能 3：Loss 平整度統計（給圖表旁的數字摘要用）
# ============================================================

def flatness_stats(values) -> Optional[dict]:
    vals = [float(v) for v in values if v is not None and pd.notna(v)]
    if not vals:
        return None
    return {
        "min": min(vals),
        "max": max(vals),
        "ripple": max(vals) - min(vals),
        "avg": sum(vals) / len(vals),
    }


# ============================================================
# 5. XML 解析與補償（改用 ElementTree，取代原本的正規表達式）
# ============================================================

class XmlStructureError(Exception):
    """XML 模板格式不符預期時拋出。"""


def _find_freq_container(root: ET.Element) -> Optional[ET.Element]:
    """
    找出「直接包含帶有 <Frequency> 子節點的 <Data> 元素」的那個父節點。
    這個節點就是我們要讀取/排序/插入頻點資料的容器
    （WWAN 對應 Cfg[Option=Common]/Std[Val=LTE]，WLAN 對應 Config/Table）。
    """
    for elem in root.iter():
        for child in list(elem):
            if child.tag == "Data" and child.find("Frequency") is not None:
                return elem
    return None


def _extract_existing_entries(container: ET.Element, mode_key: str) -> "dict[str, dict[str, str]]":
    """把容器內現有的 <Data> 區塊讀成 {正規化頻率: {通道: 值}} 的字典。"""
    channels = all_channels(mode_key)
    entries: dict[str, dict[str, str]] = {}
    for data_el in list(container.findall("Data")):
        freq_el = data_el.find("Frequency")
        if freq_el is None or freq_el.text is None:
            continue
        norm_freq = normalize_freq(freq_el.text)
        values = {}
        for ch in channels:
            ch_el = data_el.find(ch)
            values[ch] = ch_el.text if (ch_el is not None and ch_el.text is not None) else None
        entries[norm_freq] = values
    return entries


def _merge_excel_rows(entries: dict, df: pd.DataFrame, mode_key: str) -> dict:
    """用 Excel 編輯後的資料覆蓋/新增 entries（就地修改後回傳統計資訊）。"""
    cfg = CHANNEL_CONFIG[mode_key]
    channels = all_channels(mode_key)
    updated_count = 0
    added_count = 0

    for _, row in df.iterrows():
        freq_raw = row.get("Frequency")
        if pd.isna(freq_raw) or str(freq_raw).strip() == "":
            continue
        norm_freq = normalize_freq(freq_raw)

        is_new = norm_freq not in entries
        if is_new:
            entries[norm_freq] = {ch: None for ch in channels}
            added_count += 1
        else:
            updated_count += 1

        for ch in channels:
            if ch not in df.columns:
                continue
            val = row.get(ch)
            if pd.notna(val) and str(val).strip() != "":
                entries[norm_freq][ch] = str(val).strip()

        # 補上預設值（新頻點才需要，既有頻點若沒填就沿用原本 XML 的值）
        for ch in cfg["primary"]:
            if entries[norm_freq].get(ch) is None:
                entries[norm_freq][ch] = cfg["primary_default"]
        for ch in cfg["extra"]:
            if entries[norm_freq].get(ch) is None:
                entries[norm_freq][ch] = cfg["extra_default"]

    return {"updated": updated_count, "added": added_count}


def _rebuild_container(container: ET.Element, entries: dict, mode_key: str) -> None:
    """依照頻率大小排序，清空容器內舊的 <Data>，重新建立子節點。"""
    channels = all_channels(mode_key)

    # 清掉舊的 Data 節點（保留容器上其他非 Data 的子節點，例如它自己的屬性等）
    for data_el in list(container.findall("Data")):
        container.remove(data_el)

    def sort_key(item):
        norm_freq = item[0]
        try:
            return float(norm_freq)
        except (TypeError, ValueError):
            return float("inf")

    for norm_freq, values in sorted(entries.items(), key=sort_key):
        data_el = ET.SubElement(container, "Data")
        freq_el = ET.SubElement(data_el, "Frequency")
        freq_el.text = norm_freq
        for ch in channels:
            ch_el = ET.SubElement(data_el, ch)
            ch_el.text = values.get(ch) if values.get(ch) is not None else "0"


def _pretty_print(root: ET.Element) -> str:
    """輸出縮排格式化的 XML 字串（不含 <?xml ...?> 宣告，維持與原模板一致的風格）。"""
    rough = ET.tostring(root, encoding="unicode")
    parsed = minidom.parseString(rough)
    pretty = parsed.toprettyxml(indent="  ")
    # minidom 會產生 <?xml version="1.0" ?> 開頭與多餘空行，這裡去掉以貼近原始模板風格
    lines = [line for line in pretty.split("\n") if line.strip()]
    if lines and lines[0].startswith("<?xml"):
        lines = lines[1:]
    return "\n".join(lines)


def build_updated_xml(xml_content: str, mode_key: str, edited_df: pd.DataFrame) -> tuple[str, dict]:
    """
    主要進入點：讀入原始 XML 字串 + 編輯後的 DataFrame，回傳 (新的 XML 字串, 統計資訊)。
    統計資訊包含 {"updated": int, "added": int}，可用於顯示轉換結果。
    """
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        raise XmlStructureError(f"XML 格式錯誤，無法解析：{e}") from e

    container = _find_freq_container(root)
    if container is None:
        raise XmlStructureError("在 XML 中找不到含有 <Frequency> 的 <Data> 區塊，請確認模板格式是否正確。")

    entries = _extract_existing_entries(container, mode_key)
    stats = _merge_excel_rows(entries, edited_df, mode_key)
    _rebuild_container(container, entries, mode_key)

    return _pretty_print(root), stats


# ============================================================
# 6. 從既有 XML 直接讀取初始 Loss 表格
# ============================================================

def parse_loss_dataframe_from_xml(xml_content: str, mode_key: str) -> pd.DataFrame:
    """讀入一份既有的 Loss Data XML，直接解析出 Frequency + 各通道 Loss 數值，
    回傳可以當作「1. 上傳 Loss Data Excel」替代來源的 working_df 初始表格。
    """
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        raise XmlStructureError(f"XML 格式錯誤，無法解析：{e}") from e

    container = _find_freq_container(root)
    if container is None:
        raise XmlStructureError("在 XML 中找不到含有 <Frequency> 的 <Data> 區塊，請確認檔案格式是否正確。")

    entries = _extract_existing_entries(container, mode_key)
    if not entries:
        raise XmlStructureError("這份 XML 裡沒有任何 <Data><Frequency>...</Frequency></Data> 頻點資料。")

    channels = all_channels(mode_key)
    rows = []
    for norm_freq, values in entries.items():
        row = {"Frequency": norm_freq}
        for ch in channels:
            v = values.get(ch)
            try:
                row[ch] = float(v) if v is not None else None
            except (TypeError, ValueError):
                row[ch] = None
        rows.append(row)

    df = pd.DataFrame(rows)
    df["_freq_num"] = pd.to_numeric(df["Frequency"], errors="coerce")
    df = df.sort_values("_freq_num").drop(columns=["_freq_num"]).reset_index(drop=True)
    return df


# ============================================================
# 7. WWAN Band List：由 Band 清單 + Frequency List 展開出初始頻率點
# ============================================================

_BAND_TOKEN_RE = re.compile(r"[Bb]\s*\d+")


def parse_band_list_text(raw_text: str) -> list[str]:
    """從純文字（例如 "B1, B2, B12" 或每行一個）解析出 Band 清單，標準化成大寫 "B1" 這種格式。"""
    tokens = _BAND_TOKEN_RE.findall(raw_text or "")
    bands = []
    seen = set()
    for t in tokens:
        norm = "B" + re.sub(r"\s+", "", t)[1:]
        if norm not in seen:
            seen.add(norm)
            bands.append(norm)
    return bands


def parse_band_list_dataframe(df: pd.DataFrame) -> list[str]:
    """從 Excel/CSV 讀進來的 DataFrame（不限欄位名稱/欄數）攤平所有儲存格文字，解析出 Band 清單。"""
    all_text = " ".join(str(v) for v in df.values.flatten() if pd.notna(v))
    return parse_band_list_text(all_text)


def expand_band_list_to_frequencies(
    band_list: list[str], freq_list_df: pd.DataFrame, mode_key: str, default_loss: float = 40.0
) -> pd.DataFrame:
    """依 Band 清單從 Frequency List 展開出要建立的頻率點，回傳可作為 working_df 初始內容的表格
    （Frequency + 各通道 Loss，統一帶入 default_loss，之後使用者可再自行調整/套用補償）。
    """
    if mode_key != "WWAN":
        raise ValueError("目前 WWAN Band List 展開功能僅支援 WWAN 模式。")
    if not band_list:
        raise ValueError("Band List 是空的，請確認上傳的檔案內容（例如 B1, B2, B12...）。")

    normalized_bands = {b.upper() for b in band_list}
    matched = freq_list_df[freq_list_df["Band"].str.upper().isin(normalized_bands)]
    if matched.empty:
        raise ValueError("Band List 裡的 Band 在 Frequency List 中都找不到對應資料，請確認拼字是否正確。")

    channels = all_channels(mode_key)
    freqs = sorted(matched["Frequency"].dropna().unique().tolist())

    rows = []
    for f in freqs:
        row = {"Frequency": normalize_freq(f)}
        for ch in channels:
            row[ch] = default_loss
        rows.append(row)

    return pd.DataFrame(rows)


# ============================================================
# 8. 預設 Frequency List 檔案（放在程式同一個資料夾時可自動帶入，不用每次手動上傳）
# ============================================================

DEFAULT_FREQ_LIST_FILENAMES = [
    "Frequency List.xlsx",
    "Frequency_List.xlsx",
    "Frequency List.xls",
    "Frequency_List.xls",
    "frequency_list.xlsx",
    "frequency list.xlsx",
]


def find_default_frequency_list_path(base_dir: str) -> "Optional[str]":
    """在指定資料夾（通常是這支程式所在的資料夾）中尋找預設的 Frequency List 檔案，
    找到就回傳完整路徑，找不到回傳 None。
    """
    if not base_dir or not os.path.isdir(base_dir):
        return None
    for name in DEFAULT_FREQ_LIST_FILENAMES:
        candidate = os.path.join(base_dir, name)
        if os.path.isfile(candidate):
            return candidate

    # 容錯：不分大小寫比對資料夾內實際檔名
    try:
        actual_files = {f.lower(): f for f in os.listdir(base_dir)}
    except OSError:
        return None
    for name in DEFAULT_FREQ_LIST_FILENAMES:
        match = actual_files.get(name.lower())
        if match:
            return os.path.join(base_dir, match)
    return None



# ============================================================
# WLAN Band List / Frequency List / Compensation helpers
# ============================================================
_WLAN_BAND_TOKEN_RE = re.compile(r"(?i)\b(2\.4G|5G|6G)\s*[,：:;\-]?\s*(\d+)\b")
_WLAN_CHANNEL_FREQ_RE = re.compile(r"^\s*(\d+)\s*\(([-+]?\d+(?:\.\d+)?)\)")
_WLAN_RX_LEVEL_RE = re.compile(r"@\s*([-+]?\d+(?:\.\d+)?)\s*dBm", re.IGNORECASE)
_WLAN_TXPOWER_RE = re.compile(r"Tx\s*power\s*=\s*([-+]?\d+(?:\.\d+)?)", re.IGNORECASE)


def normalize_wlan_function(val) -> str:
    s = str(val).strip().upper().replace(" ", "")
    if s in ("2.4", "2.4GHZ", "2G", "2.4G"):
        return "2.4G"
    if s in ("5", "5GHZ", "5G"):
        return "5G"
    if s in ("6", "6GHZ", "6G"):
        return "6G"
    return s


def parse_wlan_band_list_text(raw_text: str) -> list[dict]:
    items, seen = [], set()
    for m in _WLAN_BAND_TOKEN_RE.finditer(raw_text or ""):
        func = normalize_wlan_function(m.group(1))
        channel = int(m.group(2))
        key = (func, channel)
        if key not in seen:
            seen.add(key)
            items.append({"Function": func, "Channel": channel, "Label": f"{func} {channel}"})
    return items


def parse_wlan_band_list_dataframe(df: pd.DataFrame) -> list[dict]:
    all_text = "\n".join(" ".join(str(v) for v in row if pd.notna(v)) for row in df.values)
    return parse_wlan_band_list_text(all_text)


def load_wlan_frequency_list(freq_df: pd.DataFrame) -> pd.DataFrame:
    colmap = {}
    for c in freq_df.columns:
        key = str(c).strip().lower().replace(" ", "")
        if key in ("function", "band"):
            colmap[c] = "Function"
        elif key == "channel":
            colmap[c] = "Channel"
        elif "centerfrequency" in key or key.startswith("frequency"):
            colmap[c] = "Frequency"
        else:
            colmap[c] = str(c).strip()
    out = freq_df.rename(columns=colmap).copy()
    required = ["Function", "Channel", "Frequency"]
    missing = [c for c in required if c not in out.columns]
    if missing:
        raise ValueError(f"WLAN Frequency List 缺少必要欄位：{', '.join(missing)}")
    out["Function"] = out["Function"].apply(normalize_wlan_function)
    out["Channel"] = pd.to_numeric(out["Channel"], errors="coerce")
    out["Frequency"] = pd.to_numeric(out["Frequency"], errors="coerce")
    out = out.dropna(subset=["Channel", "Frequency"]).copy()
    out["Channel"] = out["Channel"].astype("int64")
    out["Frequency"] = out["Frequency"].astype("float64")
    return out[["Function", "Channel", "Frequency"]].drop_duplicates(["Function", "Channel"], keep="last").reset_index(drop=True)


def read_wlan_frequency_list_excel(file_or_path) -> pd.DataFrame:
    sheets = pd.read_excel(file_or_path, engine="openpyxl", sheet_name=None)
    last_error = None
    for _, df in sheets.items():
        try:
            return load_wlan_frequency_list(df)
        except Exception as e:
            last_error = e
    raise ValueError(f"找不到可解析的 WLAN Frequency List 工作表：{last_error}")


def expand_wlan_band_list_to_frequencies(wlan_band_list: list[dict], freq_list_df: pd.DataFrame, default_loss: float = 40.0) -> pd.DataFrame:
    if not wlan_band_list:
        raise ValueError("WLAN Band List 是空的，請確認內容例如：2.4G 1、2.4G 2、5G 36。")
    freq = freq_list_df.copy()
    rows, missing = [], []
    for item in wlan_band_list:
        func = normalize_wlan_function(item["Function"])
        channel = int(item["Channel"])
        hit = freq[(freq["Function"].apply(normalize_wlan_function) == func) & (pd.to_numeric(freq["Channel"], errors="coerce") == channel)]
        if hit.empty:
            missing.append(f"{func} {channel}")
            continue
        frequency = float(hit.iloc[0]["Frequency"])
        rows.append({"Frequency": normalize_freq(frequency), "Main1_IN": float(default_loss), "Main1_OUT": float(default_loss), "WLAN_Function": func, "WLAN_Channel": channel})
    if missing:
        raise ValueError(f"WLAN Frequency List 找不到以下指定 Channel：{', '.join(missing)}")
    return pd.DataFrame(rows).sort_values("Frequency", key=lambda x: pd.to_numeric(x, errors="coerce")).reset_index(drop=True)


def parse_wlan_result_csv(df: pd.DataFrame) -> pd.DataFrame:
    required = ["Channel", "Test Item", "Description"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"WLAN 測試結果缺少必要欄位：{', '.join(missing)}")
    rows = []
    for _, r in df.iterrows():
        m_ch = _WLAN_CHANNEL_FREQ_RE.search(str(r.get("Channel", "")).strip())
        if not m_ch:
            continue
        channel = int(m_ch.group(1)); frequency = float(m_ch.group(2))
        test_item = str(r.get("Test Item", "")); desc = str(r.get("Description", ""))
        m_rx = _WLAN_RX_LEVEL_RE.search(test_item)
        m_tx = _WLAN_TXPOWER_RE.search(desc)
        rows.append({
            "Channel": channel,
            "Frequency": frequency,
            "Measured_RX": float(m_rx.group(1)) if m_rx else None,
            "Measured_TX": float(m_tx.group(1)) if m_tx else None,
            "IsLimitValue": "limit value" in desc.lower(),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("WLAN 測試結果無法解析 Channel/Frequency，請確認 Channel 格式是否類似 1(2412)。")
    selected = []
    for _, g in out.groupby(["Channel", "Frequency"], dropna=False):
        limit = g[g["IsLimitValue"] == True]  # noqa: E712
        candidate = limit if not limit.empty else g[pd.notna(g["Measured_RX"])]
        selected.append((candidate if not candidate.empty else g).iloc[-1])
    return pd.DataFrame(selected).reset_index(drop=True)


def build_wlan_compensation_table(wlan_result_df: pd.DataFrame, freq_list_df: pd.DataFrame, working_df: pd.DataFrame, targets: "dict[str, float] | None" = None, band_list: "list[dict] | None" = None) -> pd.DataFrame:
    targets_ = dict(DEFAULT_ROLE_TARGET)
    if targets:
        targets_.update(targets)
    freq = freq_list_df.copy()
    if band_list:
        keys = {(normalize_wlan_function(x["Function"]), int(x["Channel"])) for x in band_list}
        freq = freq[freq.apply(lambda r: (normalize_wlan_function(r["Function"]), int(r["Channel"])) in keys, axis=1)].copy()
        target_freqs = {normalize_freq(v) for v in freq["Frequency"].dropna().tolist()}
        wlan_result_df = wlan_result_df[wlan_result_df["Frequency"].apply(normalize_freq).isin(target_freqs)].copy()
    if freq.empty:
        raise ValueError("WLAN Band List 在 WLAN Frequency List 中找不到對應資料。")
    if wlan_result_df.empty:
        raise ValueError("WLAN Band List 對應的頻率在 WLAN 測試結果中找不到資料。")
    work = working_df.copy(); work["_norm_freq"] = work["Frequency"].apply(normalize_freq)
    orig_lookup = work.drop_duplicates("_norm_freq", keep="last").set_index("_norm_freq")
    def orig_loss(norm_freq, ch):
        if norm_freq not in orig_lookup.index:
            return None
        v = orig_lookup.loc[norm_freq].get(ch)
        return float(v) if pd.notna(v) else None
    rows = []
    for _, fl in freq.iterrows():
        frequency = float(fl["Frequency"]); channel = int(fl["Channel"]); function = normalize_wlan_function(fl["Function"]); norm_freq = normalize_freq(frequency)
        hit = wlan_result_df[wlan_result_df["Frequency"].apply(normalize_freq) == norm_freq]
        meas_rx = hit.iloc[-1].get("Measured_RX") if not hit.empty else None
        meas_tx = hit.iloc[-1].get("Measured_TX") if not hit.empty else None
        row = {"Band": function, "Channel": channel, "Frequency": frequency, "Role": "RX/TX", "Overlap": False, "Selected": True}
        in_ol = orig_loss(norm_freq, "Main1_IN")
        if pd.notna(meas_rx):
            diff = compute_diff("rx", targets_["rx"], float(meas_rx)); comp = compute_compensated_loss(in_ol, diff) if in_ol is not None else None
            row.update({"Main1_IN_OriginalLoss": in_ol, "Main1_IN_Target": targets_["rx"], "Main1_IN_Measured": float(meas_rx), "Main1_IN_Diff": diff, "Main1_IN_CompLoss": comp})
        else:
            row.update({"Main1_IN_OriginalLoss": in_ol, "Main1_IN_Target": targets_["rx"], "Main1_IN_Measured": None, "Main1_IN_Diff": None, "Main1_IN_CompLoss": None})
        out_ol = orig_loss(norm_freq, "Main1_OUT")
        if pd.notna(meas_tx):
            diff = compute_diff("tx", targets_["tx"], float(meas_tx)); comp = compute_compensated_loss(out_ol, diff) if out_ol is not None else None
            row.update({"Main1_OUT_OriginalLoss": out_ol, "Main1_OUT_Target": targets_["tx"], "Main1_OUT_Measured": float(meas_tx), "Main1_OUT_Diff": diff, "Main1_OUT_CompLoss": comp})
        else:
            row.update({"Main1_OUT_OriginalLoss": out_ol, "Main1_OUT_Target": targets_["tx"], "Main1_OUT_Measured": None, "Main1_OUT_Diff": None, "Main1_OUT_CompLoss": None})
        rows.append(row)
    return pd.DataFrame(rows).sort_values("Frequency").reset_index(drop=True)


def apply_wlan_compensation_to_working(working_df: pd.DataFrame, wlan_comp_df: pd.DataFrame) -> "tuple[pd.DataFrame, int, list[str]]":
    working_df = working_df.copy()
    for ch in ("Main1_IN", "Main1_OUT"):
        if ch in working_df.columns:
            working_df[ch] = pd.to_numeric(working_df[ch], errors="coerce").astype("float64")
    if "Selected" not in wlan_comp_df.columns or len(wlan_comp_df) == 0:
        return working_df, 0, []
    selected = wlan_comp_df[wlan_comp_df["Selected"] == True]  # noqa: E712
    applied, skipped = 0, []
    for _, row in selected.iterrows():
        norm_freq = normalize_freq(row["Frequency"])
        mask = working_df["Frequency"].apply(normalize_freq) == norm_freq
        if not mask.any():
            skipped.append(f"{row.get('Band')} / CH {row.get('Channel')} / {row.get('Frequency')} MHz：目前 Loss 表格中沒有這個頻率點，已略過。")
            continue
        if pd.notna(row.get("Main1_IN_CompLoss")):
            working_df.loc[mask, "Main1_IN"] = round(float(row["Main1_IN_CompLoss"]), 3); applied += 1
        if pd.notna(row.get("Main1_OUT_CompLoss")):
            working_df.loc[mask, "Main1_OUT"] = round(float(row["Main1_OUT_CompLoss"]), 3); applied += 1
    return working_df, applied, skipped


def wlan_compensation_warnings(wlan_comp_df: pd.DataFrame) -> list[str]:
    warnings = []
    if wlan_comp_df is None or len(wlan_comp_df) == 0:
        return warnings
    for _, row in wlan_comp_df.iterrows():
        label = f"{row.get('Band')} / CH {row.get('Channel')} / {row.get('Frequency')} MHz"
        if pd.isna(row.get("Main1_IN_OriginalLoss")):
            warnings.append(f"⚠️ {label}：Loss 表格中找不到這個頻率，缺少 Main1_IN 原Loss。")
        elif pd.isna(row.get("Main1_IN_Measured")):
            warnings.append(f"⚠️ {label}：缺少 RX sensitivity Limit Value，無法計算 Main1_IN 補償。")
    return warnings
# ============================================================
# Streamlit UI - integrated in the same file
# ============================================================

import io
import os

import streamlit as st
import pandas as pd


# ===== 網頁 UI 設定 =====
st.set_page_config(page_title="Loss Data 轉換器", layout="wide")
st.title("📡 Cable Loss 轉換 & 補償程式")
st.caption("✅ 還工作啊？趕快轉一轉看股票比較實在！")

# 1. 模式選擇
mode = st.radio("選擇頻段需求對應的 Function", ["WWAN (LTE)", "WLAN (Wi-Fi)"], horizontal=True)
mode_key = "WWAN" if "WWAN" in mode else "WLAN"
cfg = CHANNEL_CONFIG[mode_key]
primary_channels = cfg["primary"]

# ============================================================
# 側欄：共用設定（Base XML 模板 / Frequency List），全部功能共用同一份
# ============================================================
with st.sidebar:
    st.header("⚙️ 共用設定")

    st.subheader("📄 Base XML 模板")
    uploaded_xml = st.file_uploader(
        "上傳 Base XML 模板（選填，不上傳將自動帶入預設模板）", type=["xml"], key="base_xml_uploader"
    )
    st.caption("轉換步驟輸出 XML 時會用這份檔案的格式當模板；若不上傳則使用內建預設模板。")

    st.markdown("---")
    st.subheader("📑 Frequency List")
    default_freq_list_path = find_default_frequency_list_path(os.path.dirname(os.path.abspath(__file__)))
    if mode_key == "WLAN":
        for _wlan_name in ["WLAN Frequency list.xlsx", "WLAN Frequency List.xlsx", "WLAN_Frequency_List.xlsx", "wlan frequency list.xlsx"]:
            _candidate = os.path.join(os.path.dirname(os.path.abspath(__file__)), _wlan_name)
            if os.path.isfile(_candidate):
                default_freq_list_path = _candidate
                break
    freq_list_upload = st.file_uploader(
        "上傳 Frequency List Excel（可覆蓋預設檔案）", type=["xlsx", "xls"], key="shared_freq_list_uploader"
    )

    freq_list_df = None
    freq_list_error = None
    freq_list_source_label = None

    if freq_list_upload is not None:
        try:
            if mode_key == "WLAN":
                freq_list_df = read_wlan_frequency_list_excel(freq_list_upload)
            else:
                freq_list_raw = pd.read_excel(freq_list_upload, engine="openpyxl")
                freq_list_df = load_frequency_list(freq_list_raw)
            freq_list_source_label = f"已上傳：{freq_list_upload.name}"
        except Exception as e:
            freq_list_error = f"無法讀取上傳的 Frequency List：{e}"
    elif default_freq_list_path is not None:
        try:
            if mode_key == "WLAN":
                freq_list_df = read_wlan_frequency_list_excel(default_freq_list_path)
            else:
                freq_list_raw = pd.read_excel(default_freq_list_path, engine="openpyxl")
                freq_list_df = load_frequency_list(freq_list_raw)
            freq_list_source_label = f"預設檔案：{os.path.basename(default_freq_list_path)}"
        except Exception as e:
            freq_list_error = f"無法讀取預設 Frequency List（{os.path.basename(default_freq_list_path)}）：{e}"

    if freq_list_error:
        st.error(f"❌ {freq_list_error}")
    if freq_list_source_label:
        st.caption(f"✅ 目前使用：{freq_list_source_label}")
    elif not freq_list_error:
        st.caption("⚠️ 尚未提供 Frequency List，「WWAN Band List」起始資料與「MVO5 補償比對」功能暫時無法使用。")

# ============================================================
# 1. 建立初始 Loss 表格
# ============================================================
st.markdown("**📂 1. 建立初始 Loss 表格**")
source_options = ["Excel 檔案", "既有 Loss Data XML"]
if mode_key == "WWAN":
    source_options.append("WWAN Band List")
elif mode_key == "WLAN":
    source_options.append("WLAN Band List")
source_mode = st.radio(
    "資料來源", source_options, horizontal=True, key="loss_source_mode",
    help="Excel：上傳已整理好的 Loss Data 表格。既有 XML：直接讀取一份現成的 Loss Data XML 帶出頻率與 Loss。"
    "Band List：WWAN 可輸入 B1, B2；WLAN 可輸入 2.4G 1、2.4G 2、5G 36，系統會依側欄的 Frequency List 自動展開成頻率點。",
)

uploaded_excel = None
uploaded_loss_xml = None
band_list_file = None
band_list_default_loss = 40.0
band_list_from_step1 = None  # 若步驟1用「WWAN Band List」建表，這裡會存下解析出的 Band 清單，供步驟4 MVO5 篩選重複使用

if source_mode == "Excel 檔案":
    uploaded_excel = st.file_uploader("📂 上傳 Loss Data Excel", type=["xlsx", "xls"], key="loss_excel_uploader")
elif source_mode == "既有 Loss Data XML":
    uploaded_loss_xml = st.file_uploader(
        "📄 上傳既有 Loss Data XML（直接讀取 Frequency + Loss 數值）", type=["xml"], key="loss_xml_uploader"
    )
else:
    if freq_list_df is None:
        st.warning("⚠️ 請先在左側「⚙️ 共用設定」提供 Frequency List，才能使用 WWAN Band List 展開頻率點。")
    band_list_file = st.file_uploader(
        "📋 上傳 Band List（WWAN: B1, B2；WLAN: 2.4G 1, 2.4G 2, 5G 36）",
        type=["txt", "csv", "xlsx", "xls"],
        key="band_list_uploader",
    )
    band_list_default_loss = st.number_input(
        "新頻點的預設 Loss 值", value=40.0, step=0.5, format="%.2f", key="band_list_default_loss"
    )

show_extra_channels = st.checkbox(
    f"🔧 顯示其他通道欄位（{', '.join(cfg['extra'])}）",
    value=False,
    help="預設只顯示主要通道。若你的 Loss Data 也包含 Main2 / AUX 等其他通道，勾選後可在表格中一併編輯。",
)

has_source = (
    (source_mode == "Excel 檔案" and uploaded_excel is not None)
    or (source_mode == "既有 Loss Data XML" and uploaded_loss_xml is not None)
    or (source_mode in ("WWAN Band List", "WLAN Band List") and band_list_file is not None and freq_list_df is not None)
)
if not has_source:
    st.stop()

# ============================================================
# 讀取來源資料、初始化 / 重置 working_df（存在 session_state 裡才能跨按鈕互動）
# ============================================================
if source_mode == "Excel 檔案":
    try:
        raw_df = pd.read_excel(uploaded_excel, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ 無法讀取 Excel 檔案：{e}")
        st.stop()
    source_key = ("excel", uploaded_excel.file_id)

elif source_mode == "既有 Loss Data XML":
    try:
        try:
            xml_text = uploaded_loss_xml.read().decode("utf-8")
        except UnicodeDecodeError:
            uploaded_loss_xml.seek(0)
            xml_text = uploaded_loss_xml.read().decode("big5")
        raw_df = parse_loss_dataframe_from_xml(xml_text, mode_key)
    except XmlStructureError as e:
        st.error(f"❌ {e}")
        st.stop()
    except Exception as e:
        st.error(f"❌ 無法讀取 Loss Data XML：{e}")
        st.stop()
    source_key = ("xml", uploaded_loss_xml.file_id)

else:  # WWAN / WLAN Band List
    try:
        bl_name = band_list_file.name.lower()
        if source_mode == "WLAN Band List":
            if bl_name.endswith(".txt"):
                band_text = band_list_file.read().decode("utf-8", errors="ignore")
                wlan_band_list = parse_wlan_band_list_text(band_text)
            elif bl_name.endswith(".csv"):
                band_df_raw = pd.read_csv(band_list_file, header=None)
                wlan_band_list = parse_wlan_band_list_dataframe(band_df_raw)
            else:
                band_df_raw = pd.read_excel(band_list_file, engine="openpyxl", header=None)
                wlan_band_list = parse_wlan_band_list_dataframe(band_df_raw)

            raw_df = expand_wlan_band_list_to_frequencies(
                wlan_band_list, freq_list_df, default_loss=band_list_default_loss
            )
            band_list_from_step1 = wlan_band_list
            st.success(f"✅ 依 WLAN Band List（{', '.join([x['Label'] for x in wlan_band_list])}）展開出 {len(raw_df)} 個頻率點。")
            source_key = ("wlan_bandlist", band_list_file.file_id, freq_list_source_label, band_list_default_loss)
        else:
            if bl_name.endswith(".txt"):
                band_text = band_list_file.read().decode("utf-8", errors="ignore")
                band_list = parse_band_list_text(band_text)
            elif bl_name.endswith(".csv"):
                band_df_raw = pd.read_csv(band_list_file, header=None)
                band_list = parse_band_list_dataframe(band_df_raw)
            else:
                band_df_raw = pd.read_excel(band_list_file, engine="openpyxl", header=None)
                band_list = parse_band_list_dataframe(band_df_raw)

            raw_df = expand_band_list_to_frequencies(
                band_list, freq_list_df, mode_key, default_loss=band_list_default_loss
            )
            band_list_from_step1 = band_list
            st.success(f"✅ 依 Band List（{', '.join(band_list)}）展開出 {len(raw_df)} 個頻率點。")
            source_key = ("bandlist", band_list_file.file_id, freq_list_source_label, band_list_default_loss)
    except (ValueError, XmlStructureError) as e:
        st.error(f"❌ {e}")
        st.stop()
    except Exception as e:
        st.error(f"❌ 無法處理 Band List：{e}")
        st.stop()

if "Frequency" in raw_df.columns:
    raw_df["Frequency"] = raw_df["Frequency"].apply(normalize_freq)

display_cols = ["Frequency"] + primary_channels
if show_extra_channels:
    display_cols += [c for c in cfg["extra"] if c in raw_df.columns] or cfg["extra"]

for c in display_cols:
    if c not in raw_df.columns:
        raw_df[c] = None
other_cols = [c for c in raw_df.columns if c not in display_cols]
raw_df = raw_df[display_cols + other_cols]

# 只有在「換了來源檔案 / 換了模式」時才重置 working_df，避免使用者的編輯被蓋掉
load_key = (source_mode, source_key, mode_key, show_extra_channels)
if st.session_state.get("load_key") != load_key:
    st.session_state["load_key"] = load_key
    st.session_state["working_df"] = raw_df.copy()
    st.session_state["mvo5_comp_df"] = None

st.markdown("---")
st.subheader("📝 3. 編輯 / 新增頻率點")
st.info("💡 你可以直接在下方表格最底部的空白列輸入數值來 **新增頻點**，或者修改現有數值。系統會自動按頻率大小排序！")

# ============================================================
# 功能 1：統一 Loss 數值
# ============================================================
with st.expander("🎚️ 功能 1：統一 Loss 數值（批次帶入同一個可調的數值）"):
    uc1, uc2, uc3 = st.columns([2, 2, 1])
    with uc1:
        uniform_channels = st.multiselect(
            "要套用的通道",
            options=primary_channels + (cfg["extra"] if show_extra_channels else []),
            default=primary_channels,
        )
    with uc2:
        uniform_value = st.number_input("統一 Loss 數值", value=40.0, step=0.5, format="%.2f")
    with uc3:
        st.write("")
        st.write("")
        if st.button("套用到全部頻點"):
            st.session_state["working_df"] = apply_uniform_loss(
                st.session_state["working_df"], uniform_channels, uniform_value
            )
            st.rerun()

# ============================================================
# 主要 Loss 表格（互動編輯）
# ============================================================
edited_df = st.data_editor(
    st.session_state["working_df"], num_rows="dynamic", use_container_width=True, key="main_editor"
)
st.session_state["working_df"] = edited_df

# ===== 資料驗證 =====
validation = validate_dataframe(edited_df, mode_key)
for w in validation.warnings:
    st.warning(f"⚠️ {w}")
for e in validation.errors:
    st.error(f"❌ {e}")

# ============================================================
# 功能 2b：以量產測試結果 (MVO5 CSV/Excel) + Frequency List 自動比對補償 (目前僅支援 WWAN)
# ============================================================
if mode_key == "WWAN":
    st.markdown("---")
    st.subheader("🧪 4. 以量產測試結果 (MVO5) 自動比對 Frequency List 並補償")
    st.caption(
        "可一次上傳多個測試結果檔案（CSV 或 Excel），欄位需含 Band / Channel / Result(TIS) / Description 內的 Tx power(TRP)。"
        "會用左側「⚙️ 共用設定」中的 Frequency List（Band / Channel / Frequency / TX-RX）比對，"
        "系統會依 Channel 查出對應的 Frequency 與角色，自動計算補償值。"
        "角色為 RX 只更新 Main1_DL、TX 只更新 Main1_UL、TXRX 兩者都更新。"
        "若不同 Band 對應到同一個 Frequency（重疊），預設只勾選 Band 數字較小的那筆，可在表格中手動調整。"
        "套用時只會更新「目前 Loss 表格中已存在」的頻率點，測試結果中若有目前表格沒有的頻率點，不會被自動新增，也不會出現在最終輸出的 XML 中。"
    )

    if freq_list_df is None:
        st.warning("⚠️ 請先在左側「⚙️ 共用設定」提供 Frequency List，才能使用這個功能。")

    mvo5_files = st.file_uploader(
        "📊 上傳 MVO5 測試結果檔案（Result_F，可一次上傳多個 CSV/Excel 檔）",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="mvo5_files_uploader",
    )

    with st.expander("🎯 (選填) 只比對指定 Band，縮小比對範圍"):
        st.caption(
            "若這次量產測試只涉及少數幾個 Band（例如只有 B1、B2），可指定 Band 清單，"
            "系統只會用 Frequency List／測試結果中這幾個 Band 的資料來比對補償，不用掃描全部 Band。"
        )
        mvo5_band_list_file = st.file_uploader(
            "上傳 Band List（txt/csv/xlsx，內容例如 B1, B2）",
            type=["txt", "csv", "xlsx", "xls"],
            key="mvo5_band_list_uploader",
        )
        mvo5_reuse_step1_band_list = False
        if band_list_from_step1:
            mvo5_reuse_step1_band_list = st.checkbox(
                f"沿用步驟 1 的 Band List（{', '.join(band_list_from_step1)}）",
                value=(mvo5_band_list_file is None),
                key="mvo5_reuse_band_list",
            )

    mvo5_band_list = None
    if mvo5_band_list_file is not None:
        try:
            mvo5_bl_name = mvo5_band_list_file.name.lower()
            if mvo5_bl_name.endswith(".txt"):
                mvo5_band_list = parse_band_list_text(
                    mvo5_band_list_file.read().decode("utf-8", errors="ignore")
                )
            elif mvo5_bl_name.endswith(".csv"):
                mvo5_band_list = parse_band_list_dataframe(pd.read_csv(mvo5_band_list_file, header=None))
            else:
                mvo5_band_list = parse_band_list_dataframe(
                    pd.read_excel(mvo5_band_list_file, engine="openpyxl", header=None)
                )
        except Exception as e:
            st.error(f"❌ 無法讀取 Band List 檔案：{e}")
    elif mvo5_reuse_step1_band_list:
        mvo5_band_list = band_list_from_step1

    if mvo5_band_list:
        st.caption(f"✅ 這次比對只會鎖定 Band：{', '.join(mvo5_band_list)}")

    mvo5_target_col1, mvo5_target_col2, mvo5_target_col3 = st.columns([1, 1, 1])
    with mvo5_target_col1:
        mvo5_tx_target = st.number_input(
            "TX/TRP 目標值", value=23.0, step=0.5, format="%.2f", key="mvo5_tx_target"
        )
    with mvo5_target_col2:
        mvo5_rx_target = st.number_input(
            "RX/TIS 目標值", value=-101.0, step=0.5, format="%.2f", key="mvo5_rx_target"
        )
    with mvo5_target_col3:
        st.write("")
        st.write("")
        mvo5_build_clicked = st.button(
            "🔎 比對 Frequency List 並試算補償", key="mvo5_build_btn", disabled=freq_list_df is None
        )

    if mvo5_build_clicked:
        if not mvo5_files:
            st.error("❌ 請先上傳至少一個 MVO5 測試結果檔案。")
        elif freq_list_df is None:
            st.error("❌ 請先在左側「⚙️ 共用設定」提供 Frequency List。")
        else:
            mvo5_raw_dfs = []
            read_error = False
            for f in mvo5_files:
                try:
                    if f.name.lower().endswith(".csv"):
                        try:
                            df_part = pd.read_csv(f)
                        except UnicodeDecodeError:
                            f.seek(0)
                            df_part = pd.read_csv(f, encoding="big5")
                    else:
                        df_part = pd.read_excel(f, engine="openpyxl")
                    mvo5_raw_dfs.append(df_part)
                except Exception as e:
                    st.error(f"❌ 無法讀取測試結果檔案「{f.name}」：{e}")
                    read_error = True

            mvo5_raw_df = pd.concat(mvo5_raw_dfs, ignore_index=True) if (mvo5_raw_dfs and not read_error) else None

            if mvo5_raw_df is not None:
                try:
                    mvo5_parsed = parse_mvo5_result_csv(mvo5_raw_df)
                    st.session_state["mvo5_comp_df"] = build_mvo5_compensation_table(
                        mvo5_parsed,
                        freq_list_df,
                        st.session_state["working_df"],
                        targets={"tx": mvo5_tx_target, "rx": mvo5_rx_target},
                        band_list=mvo5_band_list,
                    )
                    st.rerun()
                except ValueError as e:
                    st.error(f"❌ {e}")

    if st.session_state.get("mvo5_comp_df") is not None:
        mvo5_comp_df = st.session_state["mvo5_comp_df"]
        if len(mvo5_comp_df) == 0:
            st.warning("⚠️ 測試結果中的 Channel 都沒有在 Frequency List 中找到對應資料。")
        else:
            overlap_count = int(mvo5_comp_df["Overlap"].sum())
            if overlap_count > 0:
                st.warning(
                    f"⚠️ 偵測到 {overlap_count} 筆頻率重疊（不同 Band 對應到同一個 Frequency），"
                    "已預設勾選 Band 數字較小的那筆，可在下方表格手動調整勾選狀態。"
                )

            missing_warnings = mvo5_compensation_warnings(mvo5_comp_df)
            if missing_warnings:
                with st.expander(
                    f"⚠️ 有 {len(missing_warnings)} 筆缺少原Loss或測量值，無法計算補償（點擊展開清單）", expanded=True
                ):
                    for w in missing_warnings:
                        st.markdown(f"- {w}")

            mvo5_column_config = {
                "Band": st.column_config.TextColumn("Band", disabled=True),
                "Channel": st.column_config.NumberColumn("Channel", disabled=True, format="%d"),
                "Frequency": st.column_config.NumberColumn("Frequency (MHz)", disabled=True, format="%.2f"),
                "Role": st.column_config.TextColumn("角色 (TX/RX/TXRX)", disabled=True),
                "Main1_DL_OriginalLoss": st.column_config.NumberColumn("DL 原Loss", disabled=True, format="%.2f"),
                "Main1_DL_Target": st.column_config.NumberColumn("DL 目標值(TIS)", disabled=True, format="%.2f"),
                "Main1_DL_Measured": st.column_config.NumberColumn("DL 測量值(TIS)", disabled=True, format="%.2f"),
                "Main1_DL_Diff": st.column_config.NumberColumn("DL 差值", disabled=True, format="%.2f"),
                "Main1_DL_CompLoss": st.column_config.NumberColumn("DL 補償後Loss", disabled=True, format="%.2f"),
                "Main1_UL_OriginalLoss": st.column_config.NumberColumn("UL 原Loss", disabled=True, format="%.2f"),
                "Main1_UL_Target": st.column_config.NumberColumn("UL 目標值(TRP)", disabled=True, format="%.2f"),
                "Main1_UL_Measured": st.column_config.NumberColumn("UL 測量值(TRP)", disabled=True, format="%.2f"),
                "Main1_UL_Diff": st.column_config.NumberColumn("UL 差值", disabled=True, format="%.2f"),
                "Main1_UL_CompLoss": st.column_config.NumberColumn("UL 補償後Loss", disabled=True, format="%.2f"),
                "Overlap": st.column_config.CheckboxColumn("頻率重疊", disabled=True),
                "Selected": st.column_config.CheckboxColumn("套用", disabled=False),
            }
            display_cols_mvo5 = [
                "Band", "Channel", "Frequency", "Role",
                "Main1_DL_OriginalLoss", "Main1_DL_Target", "Main1_DL_Measured", "Main1_DL_Diff", "Main1_DL_CompLoss",
                "Main1_UL_OriginalLoss", "Main1_UL_Target", "Main1_UL_Measured", "Main1_UL_Diff", "Main1_UL_CompLoss",
                "Overlap", "Selected",
            ]
            mvo5_edited = st.data_editor(
                mvo5_comp_df[display_cols_mvo5],
                use_container_width=True,
                key="mvo5_comp_editor",
                column_config=mvo5_column_config,
                hide_index=True,
            )
            st.session_state["mvo5_comp_df"] = mvo5_edited

            mvo5_dl_col1, mvo5_dl_col2 = st.columns([1, 1])
            with mvo5_dl_col1:
                mvo5_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(mvo5_excel_buffer, engine="openpyxl") as writer:
                    mvo5_edited[display_cols_mvo5].to_excel(writer, index=False, sheet_name="補償試算表")
                st.download_button(
                    label="📥 下載補償試算表 (Excel)",
                    data=mvo5_excel_buffer.getvalue(),
                    file_name="MVO5_compensation_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="mvo5_comp_download_btn",
                )
            with mvo5_dl_col2:
                if st.button("✅ 套用勾選的補償結果到上方 Loss 表格", key="mvo5_apply_btn"):
                    new_working_mvo5, applied_mvo5, skipped_mvo5 = apply_mvo5_compensation_to_working(
                        st.session_state["working_df"], st.session_state["mvo5_comp_df"]
                    )
                    st.session_state["working_df"] = new_working_mvo5
                    st.success(f"✅ 已套用 {applied_mvo5} 個通道的補償後 Loss，請往上檢查表格。")
                    if skipped_mvo5:
                        with st.expander(
                            f"⚠️ 另有 {len(skipped_mvo5)} 筆頻率點不在目前的 Loss 表格中，已略過未套用（點擊展開清單）",
                            expanded=True,
                        ):
                            for s in skipped_mvo5:
                                st.markdown(f"- {s}")
                    st.rerun()
    else:
        st.caption("上傳測試結果檔案後，點擊「比對 Frequency List 並試算補償」開始使用。")



# ============================================================
# 功能 2c：WLAN Result_P + WLAN Frequency List + Band List 自動補償
# ============================================================
if mode_key == "WLAN":
    st.markdown("---")
    st.subheader("🧪 4. WLAN 測試結果自動比對 Frequency List 並補償")
    st.caption(
        "WLAN 模式可使用 Band List（例如 2.4G 1、2.4G 2、5G 36）縮小補償範圍，"
        "再依 WLAN Frequency List 找出 Channel 對應 Frequency。Result_P 報告會解析 Description=Limit Value 那列的 "
        "Rx sensitivity search@xx dBm 作為 RX 測量值，補償 Main1_IN；若報告 Description 有 Tx power = xx，則補償 Main1_OUT。"
    )

    if freq_list_df is None:
        st.warning("⚠️ 請先在左側「⚙️ 共用設定」提供 WLAN Frequency List，才能使用這個功能。")

    wlan_files = st.file_uploader(
        "📊 上傳 WLAN 測試結果檔案（Result_P，可一次上傳多個 CSV/Excel 檔）",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="wlan_result_files_uploader",
    )

    with st.expander("🎯 (選填) 只比對指定 WLAN Band List"):
        st.caption("若這次只要補少數 WLAN Channel，例如 2.4G 1、2.4G 2、5G 36，可指定 Band List。")
        wlan_band_list_file = st.file_uploader(
            "上傳 WLAN Band List（txt/csv/xlsx，內容例如 2.4G 1、2.4G 2、5G 36）",
            type=["txt", "csv", "xlsx", "xls"],
            key="wlan_comp_band_list_uploader",
        )
        wlan_reuse_step1_band_list = False
        if band_list_from_step1 and isinstance(band_list_from_step1, list) and band_list_from_step1 and isinstance(band_list_from_step1[0], dict):
            wlan_reuse_step1_band_list = st.checkbox(
                f"沿用步驟 1 的 WLAN Band List（{', '.join([x.get('Label', '') for x in band_list_from_step1])}）",
                value=(wlan_band_list_file is None),
                key="wlan_reuse_band_list",
            )

    wlan_band_list = None
    if wlan_band_list_file is not None:
        try:
            wlan_bl_name = wlan_band_list_file.name.lower()
            if wlan_bl_name.endswith(".txt"):
                wlan_band_list = parse_wlan_band_list_text(wlan_band_list_file.read().decode("utf-8", errors="ignore"))
            elif wlan_bl_name.endswith(".csv"):
                wlan_band_list = parse_wlan_band_list_dataframe(pd.read_csv(wlan_band_list_file, header=None))
            else:
                wlan_band_list = parse_wlan_band_list_dataframe(pd.read_excel(wlan_band_list_file, engine="openpyxl", header=None))
        except Exception as e:
            st.error(f"❌ 無法讀取 WLAN Band List 檔案：{e}")
    elif wlan_reuse_step1_band_list:
        wlan_band_list = band_list_from_step1

    if wlan_band_list:
        st.caption(f"✅ 這次 WLAN 比對只會鎖定：{', '.join([x['Label'] for x in wlan_band_list])}")

    wlan_target_col1, wlan_target_col2, wlan_target_col3 = st.columns([1, 1, 1])
    with wlan_target_col1:
        wlan_tx_target = st.number_input("WLAN TX 目標值", value=23.0, step=0.5, format="%.2f", key="wlan_tx_target")
    with wlan_target_col2:
        wlan_rx_target = st.number_input("WLAN RX 目標值", value=-101.0, step=0.5, format="%.2f", key="wlan_rx_target")
    with wlan_target_col3:
        st.write(""); st.write("")
        wlan_build_clicked = st.button("🔎 比對 WLAN Frequency List 並試算補償", key="wlan_build_btn", disabled=freq_list_df is None)

    if wlan_build_clicked:
        if not wlan_files:
            st.error("❌ 請先上傳至少一個 WLAN 測試結果檔案。")
        elif freq_list_df is None:
            st.error("❌ 請先在左側「⚙️ 共用設定」提供 WLAN Frequency List。")
        else:
            wlan_raw_dfs, read_error = [], False
            for f in wlan_files:
                try:
                    if f.name.lower().endswith(".csv"):
                        try:
                            df_part = pd.read_csv(f)
                        except UnicodeDecodeError:
                            f.seek(0)
                            df_part = pd.read_csv(f, encoding="big5")
                    else:
                        df_part = pd.read_excel(f, engine="openpyxl")
                    wlan_raw_dfs.append(df_part)
                except Exception as e:
                    st.error(f"❌ 無法讀取 WLAN 測試結果「{f.name}」：{e}")
                    read_error = True
            wlan_raw_df = pd.concat(wlan_raw_dfs, ignore_index=True) if (wlan_raw_dfs and not read_error) else None
            if wlan_raw_df is not None:
                try:
                    wlan_parsed = parse_wlan_result_csv(wlan_raw_df)
                    st.session_state["wlan_comp_df"] = build_wlan_compensation_table(
                        wlan_parsed,
                        freq_list_df,
                        st.session_state["working_df"],
                        targets={"tx": wlan_tx_target, "rx": wlan_rx_target},
                        band_list=wlan_band_list,
                    )
                    st.rerun()
                except ValueError as e:
                    st.error(f"❌ {e}")

    if st.session_state.get("wlan_comp_df") is not None:
        wlan_comp_df = st.session_state["wlan_comp_df"]
        if len(wlan_comp_df) == 0:
            st.warning("⚠️ WLAN 測試結果沒有在 Frequency List 中找到對應資料。")
        else:
            missing_warnings = wlan_compensation_warnings(wlan_comp_df)
            if missing_warnings:
                with st.expander(f"⚠️ 有 {len(missing_warnings)} 筆缺少原Loss或測量值（點擊展開）", expanded=True):
                    for w in missing_warnings:
                        st.markdown(f"- {w}")

            wlan_column_config = {
                "Band": st.column_config.TextColumn("Band", disabled=True),
                "Channel": st.column_config.NumberColumn("Channel", disabled=True, format="%d"),
                "Frequency": st.column_config.NumberColumn("Frequency (MHz)", disabled=True, format="%.2f"),
                "Role": st.column_config.TextColumn("角色", disabled=True),
                "Main1_IN_OriginalLoss": st.column_config.NumberColumn("IN 原Loss", disabled=True, format="%.2f"),
                "Main1_IN_Target": st.column_config.NumberColumn("IN 目標值(RX)", disabled=True, format="%.2f"),
                "Main1_IN_Measured": st.column_config.NumberColumn("IN 測量值(RX)", disabled=True, format="%.2f"),
                "Main1_IN_Diff": st.column_config.NumberColumn("IN 差值", disabled=True, format="%.2f"),
                "Main1_IN_CompLoss": st.column_config.NumberColumn("IN 補償後Loss", disabled=True, format="%.2f"),
                "Main1_OUT_OriginalLoss": st.column_config.NumberColumn("OUT 原Loss", disabled=True, format="%.2f"),
                "Main1_OUT_Target": st.column_config.NumberColumn("OUT 目標值(TX)", disabled=True, format="%.2f"),
                "Main1_OUT_Measured": st.column_config.NumberColumn("OUT 測量值(TX)", disabled=True, format="%.2f"),
                "Main1_OUT_Diff": st.column_config.NumberColumn("OUT 差值", disabled=True, format="%.2f"),
                "Main1_OUT_CompLoss": st.column_config.NumberColumn("OUT 補償後Loss", disabled=True, format="%.2f"),
                "Overlap": st.column_config.CheckboxColumn("頻率重疊", disabled=True),
                "Selected": st.column_config.CheckboxColumn("套用", disabled=False),
            }
            display_cols_wlan = [
                "Band", "Channel", "Frequency", "Role",
                "Main1_IN_OriginalLoss", "Main1_IN_Target", "Main1_IN_Measured", "Main1_IN_Diff", "Main1_IN_CompLoss",
                "Main1_OUT_OriginalLoss", "Main1_OUT_Target", "Main1_OUT_Measured", "Main1_OUT_Diff", "Main1_OUT_CompLoss",
                "Overlap", "Selected",
            ]
            wlan_edited = st.data_editor(
                wlan_comp_df[display_cols_wlan],
                use_container_width=True,
                key="wlan_comp_editor",
                column_config=wlan_column_config,
                hide_index=True,
            )
            st.session_state["wlan_comp_df"] = wlan_edited

            wlan_dl_col1, wlan_dl_col2 = st.columns([1, 1])
            with wlan_dl_col1:
                wlan_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(wlan_excel_buffer, engine="openpyxl") as writer:
                    wlan_edited[display_cols_wlan].to_excel(writer, index=False, sheet_name="WLAN補償試算表")
                st.download_button(
                    label="📥 下載 WLAN 補償試算表 (Excel)",
                    data=wlan_excel_buffer.getvalue(),
                    file_name="WLAN_compensation_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="wlan_comp_download_btn",
                )
            with wlan_dl_col2:
                if st.button("✅ 套用勾選的 WLAN 補償結果到上方 Loss 表格", key="wlan_apply_btn"):
                    new_working_wlan, applied_wlan, skipped_wlan = apply_wlan_compensation_to_working(
                        st.session_state["working_df"], st.session_state["wlan_comp_df"]
                    )
                    st.session_state["working_df"] = new_working_wlan
                    st.success(f"✅ 已套用 {applied_wlan} 個通道的 WLAN 補償後 Loss，請往上檢查表格。")
                    if skipped_wlan:
                        with st.expander(f"⚠️ 另有 {len(skipped_wlan)} 筆頻率點不在目前 Loss 表格中，已略過", expanded=True):
                            for ss in skipped_wlan:
                                st.markdown(f"- {ss}")
                    st.rerun()
    else:
        st.caption("上傳 WLAN 測試結果後，點擊「比對 WLAN Frequency List 並試算補償」開始使用。")

# ============================================================
# 功能 3：Loss 曲線 / 平整度檢查
# ============================================================
st.markdown("---")
st.subheader("📈 5. Loss 曲線圖（依頻率由低到高顯示補償後 Loss）")

chart_source_label = "目前 Loss 表格"
chart_rows = []
comp_state_key = "mvo5_comp_df" if mode_key == "WWAN" else "wlan_comp_df"
comp_df_for_chart = st.session_state.get(comp_state_key)
combined_col = " + ".join(primary_channels)

if comp_df_for_chart is not None and len(comp_df_for_chart) > 0:
    comp_for_chart = comp_df_for_chart.copy()
    if "Selected" in comp_for_chart.columns:
        comp_for_chart = comp_for_chart[comp_for_chart["Selected"] == True]  # noqa: E712

    for _, r in comp_for_chart.iterrows():
        freq = pd.to_numeric(r.get("Frequency"), errors="coerce")
        if pd.isna(freq):
            continue
        row = {"Frequency": normalize_freq(freq), "_freq_num": float(freq)}
        combined_values = []
        has_value = False

        for ch in primary_channels:
            comp_col = f"{ch}_CompLoss"
            comp_val = pd.to_numeric(r.get(comp_col), errors="coerce")
            if pd.notna(comp_val):
                row[ch] = float(comp_val)
                combined_values.append(float(comp_val))
                has_value = True

        if combined_values:
            row[combined_col] = sum(combined_values) / len(combined_values)
        if has_value:
            chart_rows.append(row)

if chart_rows:
    chart_source = pd.DataFrame(chart_rows)
    chart_source = chart_source.sort_values("_freq_num").drop_duplicates("_freq_num", keep="last")
    chart_source_label = "補償試算表（已勾選資料的補償後 Loss）"
else:
    chart_source = st.session_state["working_df"].copy()
    chart_source["_freq_num"] = pd.to_numeric(chart_source["Frequency"], errors="coerce")
    chart_source = chart_source.dropna(subset=["_freq_num"])

chart_source = chart_source.dropna(subset=["_freq_num"]).sort_values("_freq_num")

if combined_col not in chart_source.columns:
    vals = []
    for ch in primary_channels:
        if ch in chart_source.columns:
            vals.append(pd.to_numeric(chart_source[ch], errors="coerce"))
    if vals:
        chart_source[combined_col] = pd.concat(vals, axis=1).mean(axis=1, skipna=True)

st.caption(
    f"目前圖表資料來源：{chart_source_label}；X 軸已依 Frequency MHz 由低到高排序。"
    f"合併線 {combined_col} 會把主要通道補償後 Loss 依頻率整合成同一條曲線。"
)

plot_channel_options = [c for c in (primary_channels + [combined_col] + (cfg["extra"] if show_extra_channels else [])) if c in chart_source.columns]
plot_channel_default = [c for c in primary_channels if c in plot_channel_options]
if combined_col in plot_channel_options:
    plot_channel_default.append(combined_col)

plot_channels = st.multiselect(
    "要畫在圖上的通道",
    options=plot_channel_options,
    default=plot_channel_default,
)

if plot_channels and len(chart_source) > 0:
    chart_df = chart_source.set_index("_freq_num")[plot_channels].apply(pd.to_numeric, errors="coerce").sort_index()
    st.line_chart(chart_df)

    stat_cols = st.columns(len(plot_channels))
    for col, ch in zip(stat_cols, plot_channels):
        stats = flatness_stats(chart_df[ch].dropna().tolist())
        with col:
            if stats:
                st.metric(f"{ch} 峰對峰 Ripple", f"{stats['ripple']:.2f}")
                st.caption(f"最小 {stats['min']:.2f} / 最大 {stats['max']:.2f} / 平均 {stats['avg']:.2f}")
            else:
                st.caption(f"{ch} 尚無有效數值")
else:
    st.caption("至少選一個通道且資料中需有有效 Loss 數值才會顯示圖表。")

# ============================================================
# 轉換為 XML
# ============================================================
st.markdown("---")
st.subheader("🚀 6. 執行轉換 (產生 XML)")

convert_disabled = not validation.ok
if convert_disabled:
    st.error("請先修正上方標示的錯誤，才能執行轉換。")

if st.button("🚀 執行轉換 (產生 XML)", disabled=convert_disabled):
    if uploaded_xml is not None:
        try:
            xml_content = uploaded_xml.read().decode("utf-8")
        except UnicodeDecodeError as e:
            st.error(f"❌ 無法讀取 XML 檔案（編碼錯誤）：{e}")
            st.stop()
    else:
        xml_content = DEFAULT_WWAN_XML if mode_key == "WWAN" else DEFAULT_WLAN_XML
        st.info(f"ℹ️ 未偵測到上傳模板，已自動載入預設的 {mode_key} 模板進行轉換。")

    try:
        new_xml_content, stats = build_updated_xml(xml_content, mode_key, st.session_state["working_df"])
    except XmlStructureError as e:
        st.error(f"❌ {e}")
        st.stop()

    st.success(
        f"✅ 轉換成功！更新 {stats['updated']} 個既有頻點、新增 {stats['added']} 個頻點，"
        "並已自動按頻率大小排序。請點擊下方按鈕下載。"
    )

    output_filename = "LTEloss_updated.xml" if mode_key == "WWAN" else "WLAN_loss_updated.xml"
    st.download_button(
        label="📥 下載轉換後的 XML",
        data=new_xml_content.encode("utf-8"),
        file_name=output_filename,
        mime="application/xml",
    )

    with st.expander("🔍 預覽產生的 XML"):
        st.code(new_xml_content, language="xml")