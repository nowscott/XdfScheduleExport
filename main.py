#!/usr/bin/env python3
"""
课表爬取工具 —— 主入口

使用方法:
    python3 main.py
    python3 main.py --start 2026-07-13 --end 2026-08-31

流程:
    1. 自动打开 Chrome 浏览器，导航到新东方课表页面
    2. 你在浏览器中完成登录
    3. 输入时间范围
    4. 脚本自动调 API 批量抓取并导出 Excel
"""
import argparse
import re
import sys
from datetime import datetime

from crawler import ScheduleCrawler
from excel_exporter import export_to_excel

VERSION = "1.0.0"


def validate_date(date_str: str) -> bool:
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
        return False
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def get_date_range() -> tuple[str, str]:
    print("\n" + "=" * 60)
    print("📅 请输入查询时间范围")
    print("=" * 60)

    while True:
        start = input("\n  开始日期 (格式: YYYY-MM-DD，如 2026-07-01): ").strip()
        if validate_date(start):
            break
        print("  ❌ 格式错误，请重新输入")

    while True:
        end = input("  结束日期 (格式: YYYY-MM-DD): ").strip()
        if validate_date(end):
            break
        print("  ❌ 格式错误，请重新输入")

    if start > end:
        print("  ⚠️  开始日期晚于结束日期，已自动交换")
        start, end = end, start

    return start, end


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="抓取 XDF 课表并导出 Excel 月视图")
    parser.add_argument("--start", help="开始日期，格式 YYYY-MM-DD")
    parser.add_argument("--end", help="结束日期，格式 YYYY-MM-DD")
    parser.add_argument("--headless", action="store_true", help="无界面运行（需要已有登录状态）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    args = parser.parse_args(argv)

    if bool(args.start) != bool(args.end):
        parser.error("--start 和 --end 必须同时提供")
    for value in (args.start, args.end):
        if value and not validate_date(value):
            parser.error(f"无效日期：{value}")
    return args


def main(argv: list[str] | None = None):
    args = parse_args(argv)
    print("=" * 60)
    print(f"📚 XDF 课表爬取工具 v{VERSION}")
    print("=" * 60)

    crawler = ScheduleCrawler()

    try:
        # 1. 启动浏览器
        print("\n🔧 正在启动浏览器...")
        crawler.launch(headless=args.headless)

        # 2. 等待登录
        if not crawler.wait_for_login():
            print("\n❌ 登录失败，程序退出")
            return 1

        # 3. 输入时间范围
        if args.start and args.end:
            start_date, end_date = sorted((args.start, args.end))
        else:
            start_date, end_date = get_date_range()

        # 4. 批量抓取（自动识别有课天数）
        schedules = crawler.crawl_by_date_range(start_date, end_date)

        if not schedules:
            print("\n⚠️  该时间段没有课程数据")
            return 0

        # 5. 导出 Excel
        print("\n" + "=" * 60)
        print("📊 正在导出 Excel...")
        print("=" * 60)

        output_path = export_to_excel(
            schedules,
            start_date=start_date,
            end_date=end_date,
        )

        print("\n" + "=" * 60)
        print("🎉 完成！")
        print(f"   共抓取 {len(schedules)} 条记录")
        print(f"   文件位置: {output_path}")
        print("   工作表: 课表明细 + 按月月视图（每日固定 5 个时段）")
        print("=" * 60)

    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        return 1
    finally:
        crawler.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
