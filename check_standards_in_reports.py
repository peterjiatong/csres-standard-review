#更新报告中的标准.py
import logging
import pandas as pd
from pathlib import Path
import requests
import time

from util import (
    setup_logging, MONTH_DAY, load_existing_data,
    SRC_FILE, update_std_index, extract_from_docx, get_jar, BASE_URL, HEADERS,
    process_code, remove_duplicates, get_path_for_report_folder,
    normalize_name, save_excel_with_formatting, get_path_for_log_file,
    DEST_FILE, generate_new_standards_report_in_exist_folder

)

CURRENT = {"现行", "即将实施"}

sheet_map = [
    "有搜索结果的标准",
    "无搜索结果或搜索结果过多的标准",
    "标准无详细日期(debug用)",
    "报错(debug用)",
]

dfs = pd.read_excel(SRC_FILE, sheet_name=sheet_map)
df_has_output                    = dfs["有搜索结果的标准"]
df_no_output_or_too_much_outputs = dfs["无搜索结果或搜索结果过多的标准"]
df_date_empty                    = dfs["标准无详细日期(debug用)"]
df_err                           = dfs["报错(debug用)"]

STD_INDEX = update_std_index(df_has_output)

def related_warnings(code: str) -> str | None:
    """
    针对  /XG… 修改单 以及 “E” 英文版 的补充提醒  
    - 返回 None 表示没有额外提醒；否则返回一段告警文字
    """
    # 1️⃣  英文版 (…E)
    if not code.endswith("E"):
        eng_code = f"{code}E"
        if eng_code in STD_INDEX:                 # 英文版存在
            return f"(发现英文版 {eng_code})"

    # 2️⃣  修改单  (/XGn-yyyy)
    base, _, tail = code.partition("/XG")
    if _ == "":                                   # 传入的不是“修改单”本身
        # 查找所有同基准的 /XG
        mods = sorted(k for k in STD_INDEX if k.startswith(f"{base}/XG"))
        if mods:
            return f"(存在 {len(mods)} 个修改单：{', '.join(mods)})"
    else:                                         # 传入的是某个修改单
        try:
            cur_idx = int(tail.split("-")[0])     # /XG1-2022 → 1
        except Exception:
            cur_idx = -1
        higher = sorted(
            k for k in STD_INDEX
            if k.startswith(f"{base}/XG")
            and int(k.split("/XG")[1].split("-")[0]) > cur_idx
        )
        if higher:
            return f"(存在更新的序号修改单：{', '.join(higher)})"
        else:
            return "（已是最新修改单）"

    return None

def check_one(code: str, name: str):
    """返回 (ok?, message)"""
    warn = related_warnings(code)
    if code not in STD_INDEX:
        return "no_exist", "标准库未收录（标准编号有误或不存在）"
    
    # Get status, std_name, and replacement info from STD_INDEX
    status, std_name, replacement_info = STD_INDEX[code]
    
    if status not in CURRENT:
        # Add replacement info to status_wrong message if available
        status_msg = f"状态异常（{status}）"
        if replacement_info and replacement_info.strip():
            status_msg += f" | 替代情况：{replacement_info}"
        return "status_wrong", status_msg
    
    if warn:
        if normalize_name(name) != normalize_name(std_name):
            return "name_wrong", f"{warn} | 名称不符 "
        return "ok", f"OK；{warn}"
    
    if normalize_name(name) != normalize_name(std_name):
            return "name_wrong", "名称不符"
    return "ok", "OK"

def main():
    global STD_INDEX, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err
    setup_logging("check_report_log.txt")
    logging.debug("--" * 30)
    logging.info(f"更新数据库.py - 程序开始运行 - {MONTH_DAY}")
    logging.debug("--" * 30)

    code_ok, code_err, known_codes = load_existing_data()
    codes = []
    code_to_process = []
    for docx in sorted(Path("reports").glob("*.docx")):
        hits = extract_from_docx(docx)
        if not hits:
            continue

        for _, (_, code, _) in enumerate(hits, 1):
            if code not in codes:
                codes.append(code)
                if code not in STD_INDEX:
                    code_to_process.append(code)

    logging.debug(f"待处理标准代码长度：{len(code_to_process)}")

    jar = get_jar()

    session = requests.Session()
    session.cookies.update(jar)
    logging.info("已更新cookie jar")

    try:
        session.get(BASE_URL, headers=HEADERS, timeout=(5, 15))
        logging.info("成功访问基础URL")
    except requests.ReadTimeout:
        logging.error("请求超时, 程序结束，请检查网络连接或目标网站状态")
        print("请求超时, 程序结束，请检查网络连接或目标网站状态")
        return

    for i, code in enumerate(code_to_process, 1):
        logging.info(f"更新“有搜索结果的标准”表[{i}/{len(code_to_process)}]")
        print(f"正在尝试更新标准 {i}/{len(code_to_process)}: {code}")
        process_code(code, session, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err, False)
        time.sleep(1.0)
    

    session.close()
    df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err = remove_duplicates(
        df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err
    )

    save_excel_with_formatting(DEST_FILE, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err)
    logging.info(f"⚙️  已经保存标准库至{DEST_FILE}")
    excel_log_path = get_path_for_log_file("log_excel", "standard_details.xlsx")
    save_excel_with_formatting(excel_log_path, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err)
    logging.info(f"⚙️  额外保存日志文件: {excel_log_path}")
    print("标准库更新完成")
    print("已保存标准库，并保存了日志标准库")
    print("开始检查报告")

    STD_INDEX = update_std_index(df_has_output)

    out_txt = get_path_for_report_folder("检查报告中的标准.py的运行结果", "标准检查报告.txt")
    with out_txt.open("w", encoding="utf-8") as log_f:
        for docx in sorted(Path("reports").glob("*.docx")):
            logging.info("-" * 50)
            print(f"正在检查报告：{docx.name}")
            print("-" * 50, file=log_f)

            hits = extract_from_docx(docx)
            header = f"\n📄 {docx.name} —— 共发现 {len(hits)} 条标准引用"
            logging.info(header)
            print(header, file=log_f)

            if not hits:
                continue

            for idx, (orig_code, code, name) in enumerate(hits, 1):
                status, msg = check_one(code, name)
                flag = "✅" if status == "ok" else "❌"
                if status == "no_exist":
                    line = f"{idx:>2}. {flag} {orig_code:<25} | {msg:<10} \n"
                elif status == "status_wrong":
                    line = f"{idx:>2}. {flag} {orig_code:<25} | {msg:<10} \n"
                elif status == "name_wrong":
                    try:
                        correct_name = df_has_output.loc[df_has_output["标准编号"] == orig_code, "标准名称"].iat[0]
                        line = f"{idx:>2}. {flag} {orig_code:<25} | {msg:<10} \t (正确名称应为：{correct_name}) \n"
                    except:
                        line = f"{idx:>2}. {flag} {orig_code:<25} | (本条标准问题需手动排查）\n"
                else:
                    line = f"{idx:>2}. {flag} {orig_code:<25} | {msg:<10} \n"
                logging.debug(line)
                print(line, file=log_f)

        logging.info("-" * 50)
        print("-" * 50, file=log_f)

    generate_new_standards_report_in_exist_folder(out_txt.parent, df_has_output, known_codes)
    logging.info(f"结果已同时写入 {out_txt}")
    logging.info("程序运行完成")
    print("程序运行完成")
    logging.info("--" * 30)

if __name__ == "__main__":
    main()
