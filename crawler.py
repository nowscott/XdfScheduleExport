"""Playwright 爬虫核心：登录后复用浏览器会话调用课表 API。"""
import sys
import time
from datetime import datetime, timedelta
from typing import Optional

from playwright.sync_api import sync_playwright, Page, BrowserContext, TimeoutError as PwTimeout

import config


def split_date_range_by_month(start_date: str, end_date: str) -> list[tuple[str, str]]:
    """把日期范围拆成接口可接受的逐月区间。"""
    current = datetime.strptime(start_date, "%Y-%m-%d").date()
    end = datetime.strptime(end_date, "%Y-%m-%d").date()
    ranges = []

    while current <= end:
        next_month = (current.replace(day=28) + timedelta(days=4)).replace(day=1)
        month_end = min(end, next_month - timedelta(days=1))
        ranges.append((current.isoformat(), month_end.isoformat()))
        current = next_month

    return ranges


class ScheduleCrawler:
    """
    课表爬虫

    工作流程：
      1. 启动浏览器 → 用户登录 → 进入课表页面（获得 Cookie）
      2. 直接调用后端 API 获取日历概览 + 课程详情
      3. 批量导出
    """

    def __init__(self):
        self.playwright = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.raw_schedules: list[dict] = []

    # ---- 浏览器生命周期 ----

    def launch(self, headless: bool = False) -> Page:
        """启动浏览器并导航到课表页面"""
        self.playwright = sync_playwright().start()
        config.USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(config.USER_DATA_DIR),
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
            viewport={"width": 1400, "height": 900},
        )
        self.page = self.context.new_page()

        print(f"🌐 正在打开课表页面: {config.SCHEDULE_PAGE_URL}")
        try:
            self.page.goto(config.SCHEDULE_PAGE_URL, wait_until="domcontentloaded", timeout=20000)
        except PwTimeout:
            print("⚠️  页面加载超时，继续尝试...")

        return self.page

    def wait_for_login(self, timeout: int = config.LOGIN_TIMEOUT) -> bool:
        """等待用户登录并确认课表已加载"""
        print(f"\n⏳ 请在浏览器中完成登录...（最长等待 {timeout} 秒）")
        print("   完成后脚本会自动检测，你也可以按回车跳过等待\n")

        import select
        start = time.time()
        last_state = None

        while time.time() - start < timeout:
            state = self._detect_page_state()

            if state != last_state:
                icons = {
                    "login": "🔴 仍在登录页面，请完成登录...",
                    "loading": "🟡 页面加载中，等待课表数据...",
                    "ready": "🟢 检测到课表内容！",
                    "unknown": "❓ 状态未知，请在浏览器中确认...",
                }
                print(f"  {icons.get(state, icons['unknown'])}")
                last_state = state

            if state == "ready":
                print("\n✅ 课表已加载，继续...")
                time.sleep(2)  # 确保 cookie 完全写入
                return True

            # 非阻塞检测回车
            if sys.stdin in select.select([sys.stdin], [], [], 0.5)[0]:
                sys.stdin.readline()
                print("  ✅ 手动确认继续")
                time.sleep(2)
                return True

            time.sleep(0.5)

        print("\n❌ 等待超时")
        return False

    def _detect_page_state(self) -> str:
        """通过 JS 检测页面状态"""
        try:
            result = self.page.evaluate("""() => {
                // 检测登录表单
                if (document.querySelector('input[type="password"]') ||
                    document.querySelector('button:has-text("登录")') ||
                    document.querySelector('.login-form')) {
                    return 'login';
                }
                // 检测课表内容
                const hasCalendar = document.querySelector('.el-calendar-table__row') ||
                                    document.querySelector('.fc') ||
                                    document.querySelector('[class*="schedule"]') ||
                                    document.querySelector('[class*="lesson"]');
                if (hasCalendar) return 'ready';
                // 检测加载中
                if (document.querySelector('.el-loading-mask') ||
                    document.body?.innerText?.includes('加载中')) {
                    return 'loading';
                }
                return 'unknown';
            }""")
            return result or "unknown"
        except Exception:
            return "unknown"

    def close(self):
        """关闭浏览器"""
        try:
            if self.context:
                self.context.close()
        finally:
            if self.playwright:
                self.playwright.stop()
        self.page = None
        self.context = None
        self.playwright = None
        print("👋 浏览器已关闭")

    def _ensure_page(self) -> Page:
        if self.page is None:
            raise RuntimeError("浏览器尚未启动")
        return self.page

    # ---- API 调用 ----

    def get_calendar(self, start_date: str, end_date: str) -> list[dict]:
        """
        获取日历概览（哪些天有课）
        GET /lesson/calendar/v2?startDate=&endDate=
        返回: [{day: "2026-07-04", lessonCount: 3}, ...]
        """
        resp = self._ensure_page().request.get(
            f"{config.API_BASE}/lesson/calendar/v2",
            params={"startDate": start_date, "endDate": end_date}
        )
        if not resp.ok:
            raise RuntimeError(f"日历 API HTTP {resp.status}")
        data = resp.json()
        if data.get("code") != "1":
            raise Exception(f"日历 API 异常: {data.get('msg')}")
        return data.get("data", [])

    def get_lessons(self, date_str: str) -> list[dict]:
        """
        获取某天的课程详情
        POST /lesson/list-by-date/v2
        body: {"date": "...", "lessonStatus": -1}
        """
        resp = self._ensure_page().request.post(
            f"{config.API_BASE}/lesson/list-by-date/v2",
            data={"date": date_str, "lessonStatus": -1}
        )
        if not resp.ok:
            raise RuntimeError(f"课程 API HTTP {resp.status}")
        data = resp.json()
        if data.get("code") != "1":
            raise Exception(f"课程 API 异常: {data.get('msg')}")
        return data.get("data", {}).get("lessonList") or []

    def crawl_by_date_range(self, start_date: str, end_date: str) -> list[dict]:
        """
        按时间范围批量抓取课程详情。
        策略：
          1. 先调日历 API 找出有课的天
          2. 只对有课的天调详情 API
        """
        calendar = []
        for range_start, range_end in split_date_range_by_month(start_date, end_date):
            print(f"\n📅 查询日历概览: {range_start} ~ {range_end}")
            calendar.extend(self.get_calendar(range_start, range_end))

        days_with_classes = sorted(
            (d for d in calendar if d.get("lessonCount", 0) > 0),
            key=lambda item: item["day"],
        )
        total_lessons = sum(d["lessonCount"] for d in days_with_classes)

        print(f"   有课的天数: {len(days_with_classes)} 天")
        print(f"   预计课程数: {total_lessons} 节")

        if not days_with_classes:
            print("⚠️  该时间段没有课程")
            return []

        all_lessons = []
        for i, day_info in enumerate(days_with_classes, 1):
            day = day_info["day"]
            expected = day_info["lessonCount"]
            print(f"  [{i}/{len(days_with_classes)}] {day} (预计 {expected} 节)...", end=" ")

            try:
                lessons = self.get_lessons(day)
                # 为每条记录加上日期字段
                for lesson in lessons:
                    lesson["_date"] = day
                all_lessons.extend(lessons)
                print(f"✅ {len(lessons)} 节")
            except Exception as e:
                print(f"❌ {e}")

            time.sleep(config.REQUEST_INTERVAL)

        all_lessons.sort(key=lambda item: (item.get("_date", ""), item.get("lessonStartTime", "")))
        self.raw_schedules = all_lessons
        print(f"\n📊 共抓取 {len(all_lessons)} 条课程记录")
        return all_lessons
