# -*- coding: utf-8 -*-
import os
import re
import time
import hashlib
from datetime import datetime
import pandas as pd


# ===== 统一列顺序（不存在文件时按此建表；存在则按列名对齐）=====
COL_ORDER = [
    "序号", "excel写入时间", "哈希值",
    "公历-年", "公历-月", "公历-日", "公历-时", "公历-星期",
    "农历-年", "农历-月", "农历-日", "农历-时",
    "干支-年", "干支-月", "干支-日", "干支-时",
    "旬空-年", "旬空-月", "旬空-日", "旬空-时",
    "时间1", "时间2",
    "月卦身", "世身", "八节", "神煞",
    "卦象文本", "卦象文本简介",
    "卦象名字", "本卦简称", "变卦简称",
]

# ===== 工具 =====
def md5_of_text(t: str) -> str:
    return hashlib.md5(t.encode("utf-8")).hexdigest()

def _weekday_to_num(w: str) -> str:
    if not w:
        return ""
    m = re.search(r"(星期|周)\s*([一二三四五六日天])", w)
    if not m:
        return ""
    ch = m.group(2)
    mapping = {"一":"1","二":"2","三":"3","四":"4","五":"5","六":"6","日":"7","天":"7"}
    return mapping.get(ch, "")

_CN_DIGITS = {"〇":0,"零":0,"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10,"廿":20,"卅":30,"初":0,"正":1}
def _cn_num_to_int(s: str) -> int:
    """将中文数字（含 初/十/廿/卅/正）转为整数，覆盖 1~31 常见组合。"""
    if s is None: return 0
    s = s.strip()
    s = "".join(ch for ch in s if ch in _CN_DIGITS)
    if not s:
        return 0
    if s in ("十",): return 10
    if s in ("廿",): return 20
    if s in ("卅",): return 30
    if s.startswith("初") and len(s) >= 2:
        if s[1] == "十":
            return 10
        return _CN_DIGITS.get(s[1], 0)
    if s.startswith("廿"):
        return 20 + _CN_DIGITS.get(s[1], 0) if len(s) > 1 else 20
    if s.startswith("卅"):
        return 30 + _CN_DIGITS.get(s[1], 0) if len(s) > 1 else 30
    if "十" in s:
        parts = s.split("十")
        if parts[0] == "":
            return 10 + _CN_DIGITS.get(parts[1], 0) if len(parts) > 1 else 10
        tens = _CN_DIGITS.get(parts[0], 0)
        ones = _CN_DIGITS.get(parts[1], 0) if len(parts) > 1 else 0
        return tens * 10 + ones
    return _CN_DIGITS.get(s, 0)

def _strip_paren(text: str) -> str:
    return re.sub(r"[\(（].*?[\)）]", "", text).strip()

# ===== 拆行，得到 5 条关键行（公历/农历/干支/旬空/节气行）=====
def _parse_first_block(full_text: str) -> dict:
    lines = [ln.rstrip() for ln in full_text.replace("\r\n", "\n").split("\n")]
    non_empty = [ln for ln in lines if ln.strip()]
    if len(non_empty) < 5:
        raise ValueError("文本格式不完整：关键行不足5行")
    line_g  = non_empty[0]  # 公历行（含“公历：”）
    line_n  = non_empty[1]  # 农历行（含“农历：”）
    line_gz = non_empty[2]  # 干支行
    line_xk = non_empty[3]  # 旬空行
    line_t  = non_empty[4]  # 节气时间行（“寒露…  霜降…”）
    return {"gl": line_g, "nl": line_n, "gz": line_gz, "xk": line_xk, "tline": line_t}

# ===== 逐项解析 =====
def _parse_gl(gl_line: str):
    gl = re.sub(r"^\s*公历\s*[：:]\s*", "", gl_line).strip()
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*([0-2]?\d)[:：](\d{2})\s*(.*)$", gl)
    y = mo = d = hhmm = wk = ""
    if m:
        y = m.group(1)
        mo = m.group(2)
        d = m.group(3)
        hhmm = f"{int(m.group(4)):02d}:{m.group(5)}"
        wk = _weekday_to_num(m.group(6))
    else:
        m2 = re.search(r"([0-2]?\d)[:：](\d{2})", gl)
        if m2:
            hhmm = f"{int(m2.group(1)):02d}:{m2.group(2)}"
    return y, mo, d, hhmm, wk

def _parse_nl(nl_line: str):
    nl = re.sub(r"^\s*农历\s*[：:]\s*", "", nl_line).strip()
    m_year = re.search(r"^(.*?)年", nl)
    nl_year = _strip_paren(m_year.group(1)) if m_year else ""
    m_month = re.search(r"年(.*?)月", nl)
    nl_month = _cn_num_to_int(m_month.group(1) if m_month else "")
    m_day = re.search(r"月(.*?)(?:\s|$)", nl)
    day_raw = m_day.group(1) if m_day else ""
    day_raw = re.sub(r"^[大小]", "", day_raw)
    nl_day = _cn_num_to_int(day_raw)
    toks = nl.split()
    nl_time = toks[-1].strip() if toks else ""
    return nl_year, nl_month, nl_day, nl_time

def _parse_gz(gz_line: str):
    gz = re.sub(r"^\s*干支\s*[：:]\s*", "", gz_line).strip()
    parts = [p for p in re.split(r"\s+", gz) if p]
    while len(parts) < 4: parts.append("")
    return parts[0], parts[1], parts[2], parts[3]

def _parse_xk(xk_line: str):
    xk = re.sub(r"^\s*旬空\s*[：:]\s*", "", xk_line).strip()
    parts = [p for p in re.split(r"\s+", xk) if p]
    while len(parts) < 4: parts.append("")
    return parts[0], parts[1], parts[2], parts[3]

def _parse_time_line(tline: str):
    """
    节气时间行，例如：'寒露10月8日8:42  霜降10月23日11:52'
    要求：用 1 个以上空格分割；两侧 strip 去空格
    """
    if not tline:
        return "", ""
    parts = re.split(r"\s{1,}", tline.strip())
    # 某些界面会是两个以上空格，这个正则同样适配
    if len(parts) >= 2:
        return parts[0].strip(), parts[1].strip()
    if len(parts) == 1:
        return parts[0].strip(), ""
    return "", ""

def _parse_intro(intro_text: str):
    if not intro_text:
        return {"month_pos":"", "shi_shen":"", "ba_jie":"", "shen_sha":"", "name":"", "ben_gua":"", "bian_gua":"" , "intro_all":""}
    intro_all = intro_text.strip()
    lines = intro_all.replace("\r\n","\n").split("\n")
    first = lines[0] if lines else ""
    second = lines[1] if len(lines) > 1 else ""
    parts = [p.strip() for p in first.split("；") if p.strip()]
    name = parts[0] if len(parts) >= 1 else ""
    yueguashen = parts[1] if len(parts) >= 2 else ""
    shishen = parts[2] if len(parts) >= 3 else ""
    bajie = parts[3] if len(parts) >= 4 else ""
    yueguashen = re.sub(r"^月卦身", "", yueguashen).strip()
    shishen = re.sub(r"^世身在", "", shishen).strip()
    shensha = second.strip()
    shensha = re.sub(r"^\s*神煞\s*[：:]\s*", "", shensha)
    ben = bian = ""
    if "之" in name:
        idx = name.find("之")
        ben = name[:idx]
        bian = name[idx+1:]
    return {
        "month_pos": yueguashen, "shi_shen": shishen, "ba_jie": bajie,
        "shen_sha": shensha, "name": name, "ben_gua": ben, "bian_gua": bian,
        "intro_all": intro_all
    }

# ===== 生成行数据 =====
def build_excel_row(full_text: str, intro_text: str, write_dt: datetime) -> dict:
    blk = _parse_first_block(full_text)
    gl_y, gl_m, gl_d, gl_hhmm, gl_w = _parse_gl(blk["gl"])
    nl_y, nl_m, nl_d, nl_t = _parse_nl(blk["nl"])
    gz_y, gz_m, gz_d, gz_t = _parse_gz(blk["gz"])
    xk_y, xk_m, xk_d, xk_t = _parse_xk(blk["xk"])
    time1, time2 = _parse_time_line(blk.get("tline", ""))

    intro = _parse_intro(intro_text or "")

    row = {
        "序号": "",
        "excel写入时间": write_dt.strftime("%Y-%m-%d %H:%M:%S"),
        "哈希值": md5_of_text(full_text),

        "公历-年": gl_y, "公历-月": gl_m, "公历-日": gl_d, "公历-时": gl_hhmm, "公历-星期": gl_w,
        "农历-年": nl_y, "农历-月": nl_m, "农历-日": nl_d, "农历-时": nl_t,

        "干支-年": gz_y, "干支-月": gz_m, "干支-日": gz_d, "干支-时": gz_t,
        "旬空-年": xk_y, "旬空-月": xk_m, "旬空-日": xk_d, "旬空-时": xk_t,

        "时间1": time1, "时间2": time2,

        "月卦身": intro["month_pos"], "世身": intro["shi_shen"], "八节": intro["ba_jie"], "神煞": intro["shen_sha"],
        "卦象文本": full_text,
        "卦象文本简介": intro["intro_all"],
        "卦象名字": intro["name"], "本卦简称": intro["ben_gua"], "变卦简称": intro["bian_gua"],
    }
    return row

# ===== 保存行到 Excel（按列名匹配；不存在则创建）=====
def save_row_to_excel(row: dict, path: str, extra_params=None):
    import numpy as np
    import time

    extra_params = extra_params or []
    extra_cols = [f"参数{i}" for i in range(1, len(extra_params) + 1)]

    # === 1) 读取或建空表（全部按字符串） ===
    if os.path.exists(path):
        df_old = pd.read_excel(path, dtype=str)
    else:
        df_old = pd.DataFrame(columns=COL_ORDER)

    # === 2) 仅与最后一行比对是否重复（按哈希） ===
    if not df_old.empty:
        last_row = df_old.iloc[-1]
        last_hash = str(last_row.get("哈希值", "")).strip()
        new_hash = str(row.get("哈希值", "")).strip()
        if last_hash and new_hash and last_hash == new_hash:
            print("（与上一条内容相同，跳过写入）")
            return False

    # === 3) 补齐固定列 ===
    for col in COL_ORDER:
        if col not in df_old.columns:
            df_old[col] = pd.Series(dtype=object)

    # === 4) 补齐“参数1..参数N”动态列 ===
    for col in extra_cols:
        if col not in df_old.columns:
            df_old[col] = ""

    # === 5) 写入序号 ===
    existing_rows = len(df_old)
    row = dict(row)
    row["序号"] = existing_rows + 1

    # === 6) 组装新行（固定列 + 动态参数列）===
    new_row = {c: row.get(c, "") for c in COL_ORDER}
    for i, val in enumerate(extra_params, 1):
        new_row[f"参数{i}"] = "" if val is None else str(val)

    df_row = pd.DataFrame([new_row])

    # === 7) 合并 ===
    df_new = pd.concat([df_old, df_row], ignore_index=True)

    # === 8) 调整列顺序：固定列在前，其次是已有旧列，最后追加“参数列” ===
    # 先把所有已存在列按顺序去重拼起来
    ordered_fixed = [c for c in COL_ORDER if c in df_new.columns]
    existing = [c for c in df_new.columns if c not in ordered_fixed]
    # 把参数列放到最后（保持 参数1..参数N 顺序）
    others_no_params = [c for c in existing if not c.startswith("参数")]
    params_in_df = sorted([c for c in existing if c.startswith("参数")],
                          key=lambda s: int(s.replace("参数", "")) if s.replace("参数", "").isdigit() else 10**9)
    df_new = df_new[ordered_fixed + others_no_params + params_in_df]

    # === 9) 只对固定的“数字候选列”做转换；参数列保持字符串 ===
    numeric_candidates = [
        "序号",
        "公历-年", "公历-月", "公历-日", "公历-星期",
        "农历-月", "农历-日",
    ]
    for col in numeric_candidates:
        if col in df_new.columns:
            df_new[col] = pd.to_numeric(df_new[col], errors="coerce")

    if "序号" in df_new.columns:
        df_new["序号"] = pd.to_numeric(df_new["序号"], errors="coerce").astype("Int64")

    # === 10) 写入（带重试）===
    for i in range(3):
        try:
            df_new.to_excel(path, index=False)
            print(f"已写入：{path}")
            return True
        except PermissionError:
            print(f"⚠️ Excel 文件被占用，重试 {i + 1}/3 ...")
            time.sleep(0.5)
        except Exception as e:
            print(f"❌ 写入失败：{e}")
            return False

    print("❌ 写入失败：Excel 文件仍被占用。")
    return False


# ====== Excel 中批量插入空列（UI 按钮会调用）=====
def insert_blank_cols(path: str, x_col_1based: int, col_names: list):
    if not isinstance(x_col_1based, int) or x_col_1based < 1:
        raise ValueError("X 必须是从 1 开始的正整数。")
    if (not isinstance(col_names, list)
        or any(not isinstance(n, str) or n.strip() == "" for n in col_names)):
        raise ValueError("列名必须为非空字符串列表。")

    if os.path.exists(path):
        df = pd.read_excel(path)
    else:
        df = pd.DataFrame()

    insert_loc = max(0, min(x_col_1based - 1, len(df.columns)))
    for i, nm in enumerate(col_names):
        df.insert(loc=insert_loc + i, column=nm, value="")

    df.to_excel(path, index=False)
    print(f"已在第 {x_col_1based} 列前插入 {len(col_names)} 列：{col_names}")
