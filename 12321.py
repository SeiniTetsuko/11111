import json
import re
import traceback
from datetime import datetime, timedelta
from collections import defaultdict

import requests
from zenv import get_zdkit_env


# =============================================================================
# 日志编码兜底
# =============================================================================
import sys
import builtins

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

if not getattr(builtins.print, "_patched_flush", False):
    _original_print = builtins.print

    def print(*args, **kwargs):
        kwargs.setdefault("flush", True)
        _original_print(*args, **kwargs)

    print._patched_flush = True
    builtins.print = print


# =============================================================================
# 配置加载
# =============================================================================
zenv_obj = get_zdkit_env()
BASE_URL = zenv_obj.zdkit._http_client.config.get("url")

with open(config_file.path, "r", encoding="utf-8") as f:
    config = json.load(f)

AK = config.get("ak")
SK = config.get("sk")
ORG_GUID = config.get("org_guid")
USER_GUID = config.get("user_guid")

daily_projects = config.get("daily", [])
weekly_projects = config.get("weekly", [])

RECEIVER_GUIDS = config.get("receiver_guids", [])
WEBHOOK_URLS = config.get("webhook_urls", [])


# =============================================================================
# API 路由
# =============================================================================
ACCESS_TOKEN_ROUTE = "/api/user/platform/getAccessToken"
TREE_LIST_ROUTE = "/platform/api/main/doc/treeList"
MESSAGE_SEND_ROUTE = "/middle/server/api/msg/send"

MESSAGE_TEMPLATE_ID = "80"
PLATFORM_TYPE = "all"


# =============================================================================
# 通用工具
# =============================================================================
def get_headers_with_ak(user_guid=""):
    response = requests.post(
        url=BASE_URL + ACCESS_TOKEN_ROUTE,
        json={"ak": AK, "sk": SK},
        timeout=30,
    )
    response_json = response.json()
    if not response_json.get("data"):
        raise Exception(f"获取 AccessToken 失败: {response_json}")

    access_token = response_json["data"].get("accessToken")
    return {
        "Access-Token": access_token,
        "ak": AK,
        "X-User-GUID": user_guid or USER_GUID,
    }


def normalize_to_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [x for x in value if x]
    if isinstance(value, str):
        return [value] if value.strip() else []
    return []


# =============================================================================
# 日期工具：多格式匹配
# =============================================================================
NOW = datetime.now()
TODAY_DISPLAY = NOW.strftime("%Y-%m-%d")

def get_yesterday_info():
    yesterday = NOW - timedelta(days=1)
    date_str = yesterday.strftime("%Y-%m-%d")
    return {
        "date": date_str,
        "variants": [
            date_str,                              # 2026-06-05
            date_str.replace("-", "/"),            # 2026/06/05
            date_str.replace("-", "."),            # 2026.06.05
            date_str.replace("-", ""),             # 20260605
        ],
        "display": yesterday.strftime("%Y-%m-%d"),
    }


def get_last_week_info():
    """上周一到上周日 7 天日期。"""
    last_monday = NOW - timedelta(days=NOW.weekday() + 7)
    week_dates = [last_monday + timedelta(days=i) for i in range(7)]
    date_list = [d.strftime("%Y-%m-%d") for d in week_dates]

    # 周报标题中可能出现的变体：周数、日期范围、起止日期
    week_number = last_monday.isocalendar()[1]
    year = last_monday.strftime("%Y")
    start_str = week_dates[0].strftime("%Y-%m-%d")
    end_str = week_dates[-1].strftime("%Y-%m-%d")

    variants = [
        f"W{week_number:02d}",
        f"#W{week_number:02d}",
        f"第{week_number}周",
        start_str,
        start_str.replace("-", "/"),
        start_str.replace("-", "."),
        start_str.replace("-", ""),
    ]

    return {
        "start_date": start_str,
        "end_date": end_str,
        "date_list": date_list,
        "week_number": week_number,
        "year": year,
        "variants": variants,
        "display": f"{start_str} ~ {end_str}",
    }


def is_monday():
    return NOW.weekday() == 0


def _get_tree_node_title(node):
    for key in ("dataTitle", "title", "name", "fileName", "filename"):
        value = node.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _get_tree_node_guid(node):
    for key in ("categoryGuid", "dataGuid", "guid", "fileGuid", "id"):
        value = node.get(key)
        if value:
            return value
    return ""


def title_matches_date(title, date_variants, keyword_pattern=""):
    if not title:
        return False
    has_date = any(v in title for v in date_variants)
    if not has_date:
        return False
    if keyword_pattern:
        return bool(re.search(keyword_pattern, title, re.IGNORECASE))
    return True


# =============================================================================
# 搜索文档
# =============================================================================
def search_docs(user_guid, project_guid, folder_guid, date_variants, keyword_pattern=""):
    response = requests.post(
        BASE_URL + TREE_LIST_ROUTE,
        headers=get_headers_with_ak(user_guid),
        json={"projectGuid": project_guid, "parentGuid": folder_guid},
        timeout=60,
    )

    response_json = response.json()
    doc_list = response_json.get("data") or []

    matched = []
    for doc in doc_list:
        title = _get_tree_node_title(doc)
        guid = _get_tree_node_guid(doc)
        if not guid:
            continue
        if title_matches_date(title, date_variants, keyword_pattern):
            doc_url = f"{BASE_URL}/workspace/{guid}"
            matched.append({
                "title": title,
                "guid": guid,
                "url": doc_url,
            })

    return matched


# =============================================================================
# 构建飞书导航卡片
# =============================================================================
def _build_report_table(project_results, empty_label="暂无昨日日报", title_col="标题"):
    """构建报表表格组件，有日报的排前面。"""
    rows = []
    sorted_projects = sorted(project_results, key=lambda p: 0 if p["reports"] else 1)

    for pr in sorted_projects:
        project_name = pr["project_name"]
        reports = pr["reports"]
        if reports:
            for r in reports:
                rows.append({
                    "project": project_name,
                    "title": f"[{r['title']}]({r['url']})",
                })
        else:
            rows.append({
                "project": project_name,
                "title": empty_label,
            })

    if not rows:
        return None

    return {
        "tag": "table",
        "page_size": min(len(rows),3),
        "row_height": "low",
        "freeze_first_column": True,
        "columns": [
            {
                "name": "project",
                "display_name": "项目",
                "width": "auto",
                "data_type": "text",
            },
            {
                "name": "title",
                "display_name": title_col,
                "width": "auto",
                "data_type": "lark_md",
            },
        ],
        "rows": rows,
    }


def _get_dept_level_name(pr, level_idx, fallback="未分组部门"):
    """从 dept_path 中安全获取 L3/L4/L5 名称。"""
    dept_path = pr.get("dept_path") or []
    if len(dept_path) > level_idx and str(dept_path[level_idx]).strip():
        return str(dept_path[level_idx]).strip()
    return fallback


def _count_done(items):
    """统计已找到日报/周报的部门数量。"""
    return sum(1 for item in items if item.get("reports"))


def _build_l5_report_table(items, empty_label="暂无昨日日报", title_col="日报标题"):
    """
    构建某个 L4 下的 L5 部门表格。
    一行代表一个 L5 部门；若同一个 L5 找到多篇日报/周报，则在日报列内换行展示。
    """
    rows = []

    sorted_items = sorted(
        items,
        key=lambda p: (
            _get_dept_level_name(p, 2, p.get("project_name", "")),
            p.get("project_name", ""),
        ),
    )

    for pr in sorted_items:
        reports = pr.get("reports") or []
        l5_name = _get_dept_level_name(pr, 2, pr.get("project_name", "未命名部门"))

        if reports:
            report_links = "\n".join(f"[{r['title']}]({r['url']})" for r in reports)
            rows.append({
                "dept": l5_name,
                "status": "✅ 已填写",
                "title": report_links,
            })
        else:
            rows.append({
                "dept": l5_name,
                "status": "⚠️ 未填写",
                "title": empty_label,
            })

    if not rows:
        return None

    return {
        "tag": "table",
        "page_size": min(len(rows), 8),
        "row_height": "low",
        "freeze_first_column": True,
        "header_style": {
            "background_style": "grey",
            "bold": True,
        },
        "columns": [
            {
                "name": "dept",
                "display_name": "L5 部门",
                "width": "auto",
                "data_type": "text",
                "vertical_align": "top",
            },
            {
                "name": "status",
                "display_name": "状态",
                "width": "auto",
                "data_type": "text",
                "vertical_align": "top",
            },
            {
                "name": "title",
                "display_name": title_col,
                "width": "auto",
                "data_type": "lark_md",
                "vertical_align": "top",
            },
        ],
        "rows": rows,
    }


def _build_dept_collapsible_elements(dept_results, empty_label="暂无昨日日报", title_col="日报标题"):
    """
    使用折叠面板展示部门层级：
    - L3：折叠面板
    - L4：面板内小标题
    - L5：表格行，带填写状态和日报/周报链接

    默认展开存在缺失日报/周报的 L3；全部已填写的 L3 默认收起。
    """
    if not dept_results:
        return []

    grouped = defaultdict(lambda: defaultdict(list))

    for pr in dept_results:
        l3_name = _get_dept_level_name(pr, 0, "未分组部门")
        l4_name = _get_dept_level_name(pr, 1, "直属部门")
        grouped[l3_name][l4_name].append(pr)

    elements = []

    for l3_name in sorted(grouped.keys()):
        l4_map = grouped[l3_name]
        l3_items = [item for items in l4_map.values() for item in items]
        l3_total = len(l3_items)
        l3_done = _count_done(l3_items)
        l3_missing = l3_total - l3_done

        panel_elements = [{
            "tag": "markdown",
            "content": f"**汇总：{l3_done}/{l3_total} 已填写，{l3_missing} 个缺失**",
            "text_size": "normal",
        }]

        for l4_name in sorted(l4_map.keys()):
            items = l4_map[l4_name]
            l4_total = len(items)
            l4_done = _count_done(items)
            l4_missing = l4_total - l4_done

            l4_title = f"📂 **{l4_name}**｜{l4_done}/{l4_total} 已填写"
            if l4_missing:
                l4_title += f"，{l4_missing} 个缺失"

            panel_elements.append({
                "tag": "markdown",
                "content": l4_title,
                "text_size": "normal",
            })

            table = _build_l5_report_table(
                items,
                empty_label=empty_label,
                title_col=title_col,
            )
            if table:
                panel_elements.append(table)

        panel_title = f"📍 {l3_name}｜{l3_done}/{l3_total} 已填写"
        if l3_missing:
            panel_title += f"，{l3_missing} 个缺失"

        elements.append({
            "tag": "collapsible_panel",
            "expanded": l3_missing > 0,
            "header": {
                "title": {
                    "tag": "lark_md",
                    "content": f"**{panel_title}**",
                },
                "icon": {
                    "tag": "standard_icon",
                    "token": "down-small-ccm_outlined",
                    "color": "grey",
                },
                "icon_position": "right",
                "icon_expanded_angle": -180,
                "padding": "8px 12px 8px 12px",
            },
            "border": {
                "color": "grey",
                "corner_radius": "8px",
            },
            "padding": "8px 12px 12px 12px",
            "vertical_spacing": "8px",
            "elements": panel_elements,
        })

    return elements


def build_nav_card(date_info, daily_results, weekly_results=None, week_info=None):
    """
    daily_results / weekly_results:
        [{"project_name": str, "daily_type": str, "reports": [{"title": str, "url": str}]}]
    daily_type: "platform" | "dept"
    """
    platform_daily = [p for p in daily_results if p["daily_type"] == "platform"]
    dept_daily = [p for p in daily_results if p["daily_type"] == "dept"]
    total_daily = sum(len(r["reports"]) for r in daily_results)

    has_weekly = weekly_results and any(len(r["reports"]) > 0 for r in weekly_results)
    total_weekly = sum(len(r["reports"]) for r in (weekly_results or []))
    platform_weekly = [p for p in (weekly_results or []) if p.get("daily_type") == "platform"] if weekly_results else []
    dept_weekly = [p for p in (weekly_results or []) if p.get("daily_type") == "dept"] if weekly_results else []

    # 各分区是否有内容
    platform_has_daily = any(len(p["reports"]) > 0 for p in platform_daily)
    platform_has_weekly = platform_weekly and any(len(p["reports"]) > 0 for p in platform_weekly)
    dept_has_daily = any(len(p["reports"]) > 0 for p in dept_daily)
    dept_has_weekly = dept_weekly and any(len(p["reports"]) > 0 for p in dept_weekly)

    total_reports = total_daily + total_weekly
    has_any = total_reports > 0

    header_template = "turquoise" if has_any else "red"
    header_title = f"📋{TODAY_DISPLAY} 日报导航"
    if has_any:
        header_subtitle = f"共 {total_daily} 篇日报"
        if has_weekly:
            header_subtitle += f" · {total_weekly} 篇周报"
    else:
        header_subtitle = "未找到任何昨日日报"

    elements = []

    if not has_any:
        elements.append({
            "tag": "markdown",
            "content": f"⚠️ 昨日（{date_info['display']}）未找到任何日报或周报，请及时补充填写。",
        })
    else:
        # ====== 平台项目区（日报 + 周报） ======
        if platform_daily or platform_weekly:
            elements.append({
                "tag": "markdown",
                "content": "**📂 平台项目**",
                "text_size": "heading",
            })
            # 日报子区
            if platform_has_daily:
                elements.append({
                    "tag": "markdown",
                    "content": f"**日报（{date_info['display']}）**",
                    "text_size": "normal",
                })
                table = _build_report_table(platform_daily, "暂无昨日日报", title_col="日报标题")
                if table:
                    elements.append(table)
            else:
                elements.append({
                    "tag": "markdown",
                    "content": f"**日报（{date_info['display']}）**— 暂无日报",
                    "text_size": "normal",
                })
            # 周报子区
            if platform_weekly:
                elements.append({"tag": "hr"})
                week_display = week_info["display"] if week_info else ""
                if platform_has_weekly:
                    elements.append({
                        "tag": "markdown",
                        "content": f"**周报（{week_display}）**",
                        "text_size": "normal",
                    })
                    table = _build_report_table(platform_weekly, "暂无上周周报", title_col="周报标题")
                    if table:
                        elements.append(table)
                else:
                    elements.append({
                        "tag": "markdown",
                        "content": f"**周报（{week_display}）**— 暂无周报",
                        "text_size": "normal",
                    })

        # ====== 部门区（日报 + 周报） ======
        if dept_daily or dept_weekly:
            if platform_daily or platform_weekly:
                elements.append({"tag": "hr"})
            elements.append({
                "tag": "markdown",
                "content": "**🏢 部门日报**",
                "text_size": "heading",
            })
            # 日报子区
            if dept_has_daily:
                elements.append({
                    "tag": "markdown",
                    "content": f"**日报（{date_info['display']}）**",
                    "text_size": "normal",
                })
                elements.extend(
                    _build_dept_collapsible_elements(
                        dept_daily,
                        empty_label="暂无昨日日报",
                        title_col="日报标题",
                    )
                )
            else:
                elements.append({
                    "tag": "markdown",
                    "content": f"**日报（{date_info['display']}）**— 暂无日报",
                    "text_size": "normal",
                })
            # 周报子区
            if dept_weekly:
                elements.append({"tag": "hr"})
                week_display = week_info["display"] if week_info else ""
                if dept_has_weekly:
                    elements.append({
                        "tag": "markdown",
                        "content": f"**周报（{week_display}）**",
                        "text_size": "normal",
                    })
                    elements.extend(
                        _build_dept_collapsible_elements(
                            dept_weekly,
                            empty_label="暂无上周周报",
                            title_col="周报标题",
                        )
                    )
                else:
                    elements.append({
                        "tag": "markdown",
                        "content": f"**周报（{week_display}）**— 暂无周报",
                        "text_size": "normal",
                    })

        # ====== footer ======
        elements.append({"tag": "hr"})
        footer_parts = [f"{total_daily} 篇日报"]
        if has_weekly:
            footer_parts.append(f"{total_weekly} 篇周报")
        elements.append({
            "tag": "markdown",
            "content": f"💡 共 {' · '.join(footer_parts)}，点击标题可跳转查看",
            "text_size": "notation",
        })

    return {
        "schema": "2.0",
        "header": {
            "padding": "12px 8px 12px 8px",
            "template": header_template,
            "title": {"content": header_title, "tag": "plain_text"},
            "subtitle": {"content": header_subtitle, "tag": "plain_text"},
        },
        "body": {
            "vertical_spacing": "12px",
            "elements": elements,
        },
    }


# =============================================================================
# 发送消息
# =============================================================================
def send_webhook(webhook_url, card):
    response = requests.post(
        url=webhook_url,
        headers={"Content-Type": "application/json"},
        json={"msg_type": "interactive", "card": card},
        timeout=30,
    )
    return response.json()


def send_message_api(receiver_guids, title, content, sender_guid="", interactive_content=None):
    payload = {
        "template_id": MESSAGE_TEMPLATE_ID,
        "receiver_guid": receiver_guids,
        "content": content,
        "org_guid": ORG_GUID,
        "title": title,
        "platform_type": PLATFORM_TYPE,
    }
    if interactive_content is not None:
        payload["interactive_content"] = json.dumps(interactive_content, ensure_ascii=False)

    return requests.post(
        url=BASE_URL + MESSAGE_SEND_ROUTE,
        headers=get_headers_with_ak(user_guid=sender_guid),
        json=payload,
        timeout=30,
    )


def send_card(card, title, receiver_guids, webhook_urls):
    text_content = f"【{title}】已生成，请点击查看。"

    if webhook_urls:
        for idx, url in enumerate(webhook_urls, 1):
            try:
                print(f"  📢 发送群消息 (Webhook {idx}/{len(webhook_urls)})...")
                result = send_webhook(url, card)
                if result.get("code") == 0 or result.get("StatusCode") == 0:
                    print(f"  -> ✅ 群消息发送成功")
                else:
                    print(f"  -> ❌ 群消息发送失败: {result}")
            except Exception as e:
                print(f"  -> ❌ 群消息发送异常: {e}")

    if receiver_guids:
        try:
            print(f"  📩 发送个人消息给 {len(receiver_guids)} 人...")
            response = send_message_api(
                receiver_guids=receiver_guids,
                title=title,
                content=text_content,
                sender_guid=USER_GUID,
                interactive_content=card,
            )
            if response.status_code == 200 and response.json().get("data"):
                print("  -> ✅ 个人消息发送成功")
            else:
                print(f"  -> ❌ 个人消息发送失败: {response.text}")
        except Exception as e:
            print(f"  -> ❌ 个人消息发送异常: {e}")


# =============================================================================
# 主流程
# =============================================================================
date_info = get_yesterday_info()
print("=" * 60)
print(f"开始执行日报导航 | 日期: {date_info['display']} | 日报项目: {len(daily_projects)} | 周报项目: {len(weekly_projects)}")
print("=" * 60)

# 1. 搜索日报
daily_results = []

for project in daily_projects:
    project_name = project.get("project_name", "Unknown")
    project_guid = project.get("project_guid")
    parent_guid = project.get("parent_guid")
    user_guid = project.get("user_guid") or USER_GUID

    if not project_guid:
        print(f"\n⏭ 跳过项目: {project_name} (缺少 project_guid)")
        continue

    print(f"\n▶ [日报] 搜索项目: {project_name}")

    reports = search_docs(
        user_guid=user_guid,
        project_guid=project_guid,
        folder_guid=parent_guid,
        date_variants=date_info["variants"],
    )

    print(f"  找到 {len(reports)} 篇日报")
    for r in reports:
        print(f"    - {r['title']}")

    daily_results.append({
        "project_name": project_name,
        "daily_type": project.get("daily_type", "platform"),
        "dept_path": project.get("dept_path", []),
        "reports": reports,
    })

# 2. 周一搜索上周周报
weekly_results = None
week_info = None

if is_monday() and weekly_projects:
    week_info = get_last_week_info()
    weekly_results = []

    print(f"\n{'=' * 60}")
    print(f"📅 周一自动搜索上周周报 | {week_info['display']} | 周数: W{week_info['week_number']:02d}")
    print("=" * 60)

    for project in weekly_projects:
        project_name = project.get("project_name", "Unknown")
        project_guid = project.get("project_guid")
        parent_guid = project.get("parent_guid")
        user_guid = project.get("user_guid") or USER_GUID

        if not project_guid:
            print(f"\n⏭ 跳过项目: {project_name} (缺少 project_guid)")
            continue

        print(f"\n▶ [周报] 搜索项目: {project_name}")

        reports = search_docs(
            user_guid=user_guid,
            project_guid=project_guid,
            folder_guid=parent_guid,
            date_variants=week_info["variants"],
        )

        print(f"  找到 {len(reports)} 篇周报")
        for r in reports:
            print(f"    - {r['title']}")

        weekly_results.append({
            "project_name": project_name,
            "daily_type": project.get("daily_type", "platform"),
            "dept_path": project.get("dept_path", []),
            "reports": reports,
        })

# 3. 构建汇聚卡片
card = build_nav_card(date_info, daily_results, weekly_results=weekly_results, week_info=week_info)

# 4. 发送
receiver_guids = list(dict.fromkeys(RECEIVER_GUIDS))
webhook_urls = list(dict.fromkeys(WEBHOOK_URLS))

title = f"📋{TODAY_DISPLAY} 日报导航"

print("\n" + "=" * 60)
print(f"发送导航卡片 | 接收人: {len(receiver_guids)} | 群: {len(webhook_urls)}")
print("=" * 60)

send_card(card, title, receiver_guids, webhook_urls)

print("\n" + "=" * 60)
print("导航任务执行完毕")
print("=" * 60)
