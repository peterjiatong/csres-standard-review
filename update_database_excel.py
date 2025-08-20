# 更新数据库.py
import pandas as pd
import logging
import requests
import time
import os
import sys

# Import from your util.py
from util import (
    MONTH_DAY, DEST_FILE, BASE_URL, HEADERS,
    get_path_for_log_file, get_jar, process_code, setup_logging,
    load_existing_data, initialize_dataframes, remove_duplicates,
    save_excel_with_formatting, generate_new_standards_report
)

def main():
    print("程序开始运行")
    setup_logging("update_std_log.txt")
    logging.debug("--" * 30)
    logging.info(f"更新数据库.py - 程序开始运行 - {MONTH_DAY}")
    logging.debug("--" * 30)

    code_ok, code_err, known_codes = load_existing_data()
    df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err = initialize_dataframes()
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

    for i, code in enumerate(code_ok, 1):
        logging.info(f"更新“有搜索结果的标准”表[{i}/{len(code_ok)}]")
        print(f"更新“有搜索结果的标准”表[{i}/{len(code_ok)}]")
        process_code(code, session, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err, False)
        logging.info("--" * 15)
        time.sleep(1.0)

    for i, code in enumerate(code_err, 1):
        logging.info(f"更新“无搜索结果或搜索结果过多的标准”表[{i}/{len(code_err)}]")
        print(f"更新“无搜索结果或搜索结果过多的标准”表[{i}/{len(code_err)}]")
        process_code(code, session, df_has_output, df_no_output_or_too_much_outputs, df_date_empty, df_err, True)
        logging.info("--" * 15)
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

    generate_new_standards_report(df_has_output, known_codes)

    logging.info(f"\n已完成，已保存")
    logging.info(f"处理了{len(df_has_output)}个有搜索结果的标准, "
                f"{len(df_no_output_or_too_much_outputs)}个无搜索结果或搜索结果过多的标准, "
                f"{len(df_err)}个标准出错")
    
    logging.info("程序运行完成")
    print("程序运行完成")
    logging.info("--" * 30)

if __name__ == "__main__":
    main()

