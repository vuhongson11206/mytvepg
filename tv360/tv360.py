import os
import re
import json
import html
import time
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from collections import OrderedDict

import requests
from openpyxl import load_workbook


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "tv360channels.xlsx"
EPG_FILE = BASE_DIR / "tv360epg.xml"
LOG_FILE = BASE_DIR / "tv360log.txt"

DATA_SHEET = "Data"
REFERENCE_SHEET = "Tham chiếu"

CHANNEL_API_URL = "https://m.tv360.vn/tv/"
EPG_API_URL = "https://m.tv360.vn/public/v1/live/get-live-schedule?id={}"

REQUEST_TIMEOUT = 30
REQUEST_DELAY = 0.2

MAX_LOG_RUNS = 7

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json,text/html,*/*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://m.tv360.vn/tv/",
}


# ============================================================
# HTTP SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# UTILITY
# ============================================================

def now_vietnam():
    """
    Trả về thời gian Việt Nam.
    GitHub runner thường chạy UTC nên không dùng datetime.now() đơn thuần.
    """
    from zoneinfo import ZoneInfo

    return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh"))


def clean_text(value):
    """
    Chuẩn hóa text theo yêu cầu:

    - Xóa khoảng trắng trước , :
    - Thêm khoảng trắng sau , :
    - Nếu , : ở cuối thì xóa
    - Thêm khoảng trắng trước và sau -
    - Xóa khoảng trắng kép
    """

    if value is None:
        return ""

    value = str(value)

    # HTML entity nếu có
    value = html.unescape(value)

    # --------------------------------------------------------
    # Xử lý dấu , :
    # --------------------------------------------------------

    # Xóa khoảng trắng trước dấu , :
    value = re.sub(r"\s+([,:])", r"\1", value)

    # Nếu dấu , hoặc : nằm cuối chuỗi thì xóa
    value = re.sub(r"[,:]\s*$", "", value)

    # Thêm khoảng trắng sau , :
    # Chỉ khi phía sau có ký tự
    value = re.sub(r"([,:])(?=\S)", r"\1 ", value)

    # --------------------------------------------------------
    # Xử lý dấu -
    # --------------------------------------------------------

    # Xóa khoảng trắng xung quanh -
    value = re.sub(r"\s*-\s*", " - ", value)

    # --------------------------------------------------------
    # Xóa khoảng trắng kép
    # --------------------------------------------------------

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def xml_escape(value):
    """
    Escape XML.
    Dùng html.escape để xử lý &, <, >, ".
    Sau đó thay ' thành &apos;.
    """

    value = clean_text(value)

    value = (
        value
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )

    return value


def parse_time(time_string):
    """
    HH:MM
    """

    if not time_string:
        return None

    try:
        return datetime.strptime(time_string, "%H:%M")
    except Exception:
        return None


def calculate_duration(start_time, end_time):
    """
    Tính số phút từ startTime đến endTime.

    Ví dụ:
    00:00 -> 01:30 = 90
    23:30 -> 00:15 = 45
    """

    start = parse_time(start_time)
    end = parse_time(end_time)

    if not start or not end:
        return None

    dummy_date = datetime(2000, 1, 1)

    start = dummy_date.replace(
        hour=start.hour,
        minute=start.minute
    )

    end = dummy_date.replace(
        hour=end.hour,
        minute=end.minute
    )

    if end <= start:
        end += timedelta(days=1)

    minutes = int((end - start).total_seconds() / 60)

    return minutes


def format_xmltv_datetime(date_string, time_string):
    """
    2026-08-22 + 00:00
    =>
    20260822000000 +0700
    """

    if not date_string or not time_string:
        return None

    try:
        dt = datetime.strptime(
            f"{date_string} {time_string}",
            "%Y-%m-%d %H:%M"
        )

        return dt.strftime("%Y%m%d%H%M%S") + " +0700"

    except Exception:
        return None


# ============================================================
# FIND CHANNEL JSON
# ============================================================

def extract_json_objects(text):
    """
    Tìm các object JSON trong HTML/text.

    TV360 đôi khi trả HTML có embedded JSON.
    Hàm này cố gắng tìm các object có:
        id
        slug
        link
    """

    objects = []

    decoder = json.JSONDecoder()

    for match in re.finditer(r'\{', text):

        start = match.start()

        try:
            obj, end = decoder.raw_decode(text[start:])

            if isinstance(obj, dict):
                objects.append(obj)

        except Exception:
            continue

    return objects


def get_channels_from_tv360():
    """
    Lấy toàn bộ channel từ https://m.tv360.vn/tv/
    """

    print("Đang lấy danh sách kênh TV360...")

    response = session.get(
        CHANNEL_API_URL,
        timeout=REQUEST_TIMEOUT
    )

    response.raise_for_status()

    text = response.text

    channels = OrderedDict()

    # --------------------------------------------------------
    # Trường hợp response trực tiếp là JSON
    # --------------------------------------------------------

    try:
        data = response.json()

        candidate_objects = extract_channel_objects(data)

        for obj in candidate_objects:
            add_channel(channels, obj)

    except Exception:
        pass

    # --------------------------------------------------------
    # Tìm embedded JSON trong HTML
    # --------------------------------------------------------

    if not channels:
        objects = extract_json_objects(text)

        for obj in objects:

            # Chính object
            add_channel(channels, obj)

            # Các key có thể chứa list
            for value in obj.values():

                if isinstance(value, list):

                    for item in value:

                        if isinstance(item, dict):
                            add_channel(channels, item)

    # --------------------------------------------------------
    # Kiểm tra
    # --------------------------------------------------------

    if not channels:
        raise RuntimeError(
            "Không tìm thấy danh sách kênh TV360 từ "
            f"{CHANNEL_API_URL}"
        )

    result = list(channels.values())

    print(f"Tìm thấy {len(result)} kênh.")

    return result


def extract_channel_objects(data):
    """
    Tìm channel objects trong JSON response.
    """

    result = []

    def walk(obj):

        if isinstance(obj, dict):

            if (
                "id" in obj
                and "slug" in obj
                and (
                    "link" in obj
                    or "coverImage" in obj
                    or "horizontalImage" in obj
                )
            ):
                result.append(obj)

            for value in obj.values():
                walk(value)

        elif isinstance(obj, list):

            for item in obj:
                walk(item)

    walk(data)

    return result


def add_channel(channels, obj):
    """
    Thêm một channel nếu object đúng cấu trúc.
    """

    if not isinstance(obj, dict):
        return

    if "id" not in obj:
        return

    if "slug" not in obj:
        return

    if "link" not in obj:
        return

    channel_id = str(obj.get("id", "")).strip()

    if not channel_id:
        return

    name = str(obj.get("name", "")).strip()

    slug = str(obj.get("slug", "")).strip()

    link = str(obj.get("link", "")).strip()

    # Ưu tiên horizontalImage
    horizontal_image = str(
        obj.get("horizontalImage", "") or ""
    ).strip()

    # --------------------------------------------------------
    # Một số response có name đúng.
    # Nếu không có name thì tạo từ slug.
    # --------------------------------------------------------

    if not name:
        name = slug.replace("-", " ")

    channel = {
        "id": channel_id,
        "name": name,
        "slug": slug,
        "link": link,
        "horizontalImage": horizontal_image,
    }

    channels[channel_id] = channel


# ============================================================
# EXCEL
# ============================================================

def ensure_excel_file():
    """
    Kiểm tra tv360channels.xlsx.
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Không tìm thấy file:\n{EXCEL_FILE}\n\n"
            "Hãy tạo tv360channels.xlsx và sheet "
            "'Data' + 'Tham chiếu' trước."
        )


def prepare_data_sheet(ws):
    """
    Đảm bảo header Data đúng.
    """

    headers = [
        "id",
        "name",
        "slug",
        "link",
        "horizontalImage",
        "channel",
        "display-name",
    ]

    for col, header in enumerate(headers, start=1):

        # Chỉ đặt A:E nếu header chưa đúng.
        # Không đụng F:G.
        if ws.cell(1, col).value is None:
            ws.cell(1, col).value = header


def update_data_sheet(wb, channels):
    """
    Cập nhật A:E.

    Quan trọng:
    - Không xóa F:G
    - Không xóa công thức
    - Không xóa dữ liệu cũ ở F:G
    - Không đụng sheet Tham chiếu
    """

    if DATA_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(DATA_SHEET)
    else:
        ws = wb[DATA_SHEET]

    prepare_data_sheet(ws)

    # --------------------------------------------------------
    # Xóa dữ liệu A:E cũ.
    # F:G hoàn toàn không đụng tới.
    # --------------------------------------------------------

    max_row = max(ws.max_row, len(channels) + 1)

    for row in range(2, max_row + 1):

        for col in range(1, 6):
            ws.cell(row, col).value = None

    # --------------------------------------------------------
    # Ghi channel mới
    # --------------------------------------------------------

    for row_num, channel in enumerate(channels, start=2):

        ws.cell(row_num, 1).value = channel["id"]
        ws.cell(row_num, 2).value = channel["name"]
        ws.cell(row_num, 3).value = channel["slug"]
        ws.cell(row_num, 4).value = channel["link"]
        ws.cell(row_num, 5).value = channel["horizontalImage"]

    print(
        f"Đã cập nhật Data A:E với {len(channels)} kênh."
    )


# ============================================================
# REFERENCE
# ============================================================

def load_reference_mapping(wb):
    """
    Đọc mapping từ sheet 'Tham chiếu'.

    Cấu trúc:
    A = id
    B = name
    C = slug
    D = link
    E = horizontalImage
    F = channel
    G = display-name

    Python KHÔNG dựa vào VLOOKUP.

    Điều này rất quan trọng vì openpyxl không tự tính công thức.
    """

    if REFERENCE_SHEET not in wb.sheetnames:
        raise RuntimeError(
            f"Không tìm thấy sheet '{REFERENCE_SHEET}' "
            "trong tv360channels.xlsx"
        )

    ws = wb[REFERENCE_SHEET]

    mapping = {}

    for row in range(2, ws.max_row + 1):

        channel_id = ws.cell(row, 1).value

        if channel_id is None:
            continue

        channel_id = str(channel_id).strip()

        if not channel_id:
            continue

        channel = ws.cell(row, 6).value
        display_name = ws.cell(row, 7).value

        if channel is None or str(channel).strip() == "":
            continue

        if display_name is None or str(display_name).strip() == "":
            continue

        mapping[channel_id] = {
            "channel": str(channel).strip(),
            "display-name": str(display_name).strip(),
        }

    print(
        f"Đã đọc {len(mapping)} mapping từ '{REFERENCE_SHEET}'."
    )

    return mapping


# ============================================================
# EPG API
# ============================================================

def get_channel_epg(channel_id):
    """
    Lấy EPG của một channel.
    """

    url = EPG_API_URL.format(channel_id)

    try:

        response = session.get(
            url,
            timeout=REQUEST_TIMEOUT
        )

        response.raise_for_status()

        data = response.json()

        # ----------------------------------------------------
        # Tìm schedules ở nhiều cấu trúc JSON
        # ----------------------------------------------------

        schedules = find_schedules(data)

        if not schedules:
            return []

        return schedules

    except Exception as e:

        print(
            f"[EPG ERROR] id={channel_id}: {e}"
        )

        return None


def find_schedules(data):
    """
    Tìm key schedules trong JSON.
    """

    if isinstance(data, dict):

        if isinstance(data.get("schedules"), list):
            return data["schedules"]

        for value in data.values():

            result = find_schedules(value)

            if result:
                return result

    elif isinstance(data, list):

        for item in data:

            result = find_schedules(item)

            if result:
                return result

    return []


# ============================================================
# GET EPG FOR ALL CHANNELS
# ============================================================

def collect_epg(channels, reference_mapping):

    programmes = []

    channels_with_epg = []
    channels_without_epg = []

    today = now_vietnam().strftime("%Y-%m-%d")

    for index, channel in enumerate(channels, start=1):

        channel_id = str(channel["id"])

        print(
            f"[{index}/{len(channels)}] "
            f"{channel_id} - {channel['name']}"
        )

        # ----------------------------------------------------
        # Chỉ lấy EPG cho channel có mapping.
        # ----------------------------------------------------

        if channel_id not in reference_mapping:

            print(
                "  -> Chưa có mapping trong Tham chiếu."
            )

            channels_without_epg.append({
                "id": channel_id,
                "name": channel["name"],
                "reason": "Chưa có mapping trong sheet Tham chiếu",
            })

            continue

        mapping = reference_mapping[channel_id]

        schedules = get_channel_epg(channel_id)

        if schedules is None:

            channels_without_epg.append({
                "id": channel_id,
                "name": channel["name"],
                "reason": "Lỗi khi gọi API EPG",
            })

            continue

        if len(schedules) == 0:

            channels_without_epg.append({
                "id": channel_id,
                "name": channel["name"],
                "reason": "API không có schedules",
            })

            continue

        valid_count = 0

        for schedule in schedules:

            if not isinstance(schedule, dict):
                continue

            name = schedule.get("name", "")
            start_time = schedule.get("startTime", "")
            end_time = schedule.get("endTime", "")

            # ------------------------------------------------
            # Date
            # ------------------------------------------------

            date_string = (
                schedule.get("datetime")
                or schedule.get("date")
                or today
            )

            # date có thể là "Hôm nay"
            if date_string == "Hôm nay":
                date_string = today

            # ------------------------------------------------
            # Kiểm tra dữ liệu
            # ------------------------------------------------

            if not name:
                continue

            if not start_time or not end_time:
                continue

            start_xml = format_xmltv_datetime(
                date_string,
                start_time
            )

            stop_xml = format_xmltv_datetime(
                date_string,
                end_time
            )

            if not start_xml or not stop_xml:
                continue

            duration = calculate_duration(
                start_time,
                end_time
            )

            if duration is None:
                continue

            programmes.append({
                "channel": mapping["channel"],
                "display-name": mapping["display-name"],
                "start": start_xml,
                "stop": stop_xml,
                "title": clean_text(name),
                "duration": duration,
            })

            valid_count += 1

        if valid_count > 0:

            channels_with_epg.append({
                "id": channel_id,
                "name": channel["name"],
                "programmes": valid_count,
            })

        else:

            channels_without_epg.append({
                "id": channel_id,
                "name": channel["name"],
                "reason": "Có schedules nhưng không có dữ liệu hợp lệ",
            })

        time.sleep(REQUEST_DELAY)

    return (
        programmes,
        channels_with_epg,
        channels_without_epg
    )


# ============================================================
# XML
# ============================================================

def create_epg_xml(
    channels,
    reference_mapping,
    programmes
):

    print("Đang tạo tv360epg.xml...")

    # --------------------------------------------------------
    # Chỉ tạo channel XML cho những channel có mapping.
    # --------------------------------------------------------

    xml_lines = []

    xml_lines.append(
        '<?xml version="1.0" encoding="UTF-8"?>'
    )

    xml_lines.append(
        '<tv source-info-name="Ngân Phúc" '
        'source-info-url="https://epg.vercel.app/epg.xml" '
        'generator-info-name="EPG GitHub">'
    )

    # --------------------------------------------------------
    # Channel
    # --------------------------------------------------------

    added_channels = set()

    for channel in channels:

        channel_id = str(channel["id"])

        if channel_id not in reference_mapping:
            continue

        mapping = reference_mapping[channel_id]

        xml_channel = mapping["channel"]

        if xml_channel in added_channels:
            continue

        added_channels.add(xml_channel)

        display_name = xml_escape(
            mapping["display-name"]
        )

        xml_lines.append(
            f'  <channel id="{xml_escape(xml_channel)}">'
        )

        xml_lines.append(
            f'    <display-name lang="vi">'
            f'{display_name}'
            f'</display-name>'
        )

        xml_lines.append(
            '  </channel>'
        )

    # --------------------------------------------------------
    # Programme
    # --------------------------------------------------------

    programmes_sorted = sorted(
        programmes,
        key=lambda x: (
            x["channel"],
            x["start"]
        )
    )

    for program in programmes_sorted:

        channel_id = xml_escape(
            program["channel"]
        )

        start = program["start"]
        stop = program["stop"]

        title = xml_escape(
            program["title"]
        )

        duration = program["duration"]

        xml_lines.append(
            f'  <programme '
            f'start="{start}" '
            f'stop="{stop}" '
            f'channel="{channel_id}">'
        )

        xml_lines.append(
            f'    <title lang="vi">'
            f'{title}'
            f'</title>'
        )

        xml_lines.append(
            f'    <length lang="vi">'
            f'Chương trình này có thời lượng '
            f'{duration} phút'
            f'</length>'
        )

        xml_lines.append(
            '  </programme>'
        )

    xml_lines.append('</tv>')

    # --------------------------------------------------------
    # Ghi UTF-8
    # --------------------------------------------------------

    with open(
        EPG_FILE,
        "w",
        encoding="utf-8",
        newline="\n"
    ) as f:

        f.write("\n".join(xml_lines))

    print(
        f"Đã tạo {EPG_FILE}"
    )

    return len(programmes)


# ============================================================
# LOG
# ============================================================

def load_previous_log():
    """
    Đọc log cũ.

    Log được lưu theo dạng JSON block để Python dễ so sánh
    với lần chạy trước.
    """

    if not LOG_FILE.exists():
        return []

    try:

        text = LOG_FILE.read_text(
            encoding="utf-8"
        )

        blocks = []

        # ----------------------------------------------------
        # Mỗi run bắt đầu bằng:
        # ====================================================
        # RUN_JSON_START
        # {...}
        # RUN_JSON_END
        # ----------------------------------------------------

        pattern = re.compile(
            r"RUN_JSON_START\s*"
            r"(.*?)"
            r"\s*RUN_JSON_END",
            re.DOTALL
        )

        for match in pattern.finditer(text):

            try:

                data = json.loads(
                    match.group(1)
                )

                blocks.append(data)

            except Exception:
                continue

        return blocks

    except Exception:
        return []


def build_log(
    current_channels,
    channels_with_epg,
    channels_without_epg
):

    previous_runs = load_previous_log()

    previous = (
        previous_runs[0]
        if previous_runs
        else None
    )

    current_map = OrderedDict()

    for channel in current_channels:

        current_map[
            str(channel["id"])
        ] = channel["name"]

    previous_map = {}

    if previous:

        previous_map = previous.get(
            "channels",
            {}
        )

    current_ids = set(current_map)
    previous_ids = set(previous_map)

    new_ids = current_ids - previous_ids
    removed_ids = previous_ids - current_ids

    new_channels = [
        {
            "id": channel_id,
            "name": current_map[channel_id]
        }
        for channel_id in sorted(
            new_ids,
            key=lambda x: int(x)
            if x.isdigit()
            else x
        )
    ]

    removed_channels = [
        {
            "id": channel_id,
            "name": previous_map[channel_id]
        }
        for channel_id in sorted(
            removed_ids,
            key=lambda x: int(x)
            if x.isdigit()
            else x
        )
    ]

    current_count = len(current_channels)

    previous_count = (
        len(previous_map)
        if previous
        else 0
    )

    difference = (
        current_count - previous_count
    )

    # --------------------------------------------------------
    # JSON data dùng cho lần chạy tiếp theo
    # --------------------------------------------------------

    json_data = {
        "timestamp": now_vietnam().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        "channel_count": current_count,
        "channels": dict(current_map),
        "epg_channel_count": len(
            channels_with_epg
        ),
        "no_epg_channel_count": len(
            channels_without_epg
        ),
    }

    # --------------------------------------------------------
    # Text log
    # --------------------------------------------------------

    lines = []

    lines.append(
        "=" * 80
    )

    lines.append(
        f"TV360 - {json_data['timestamp']}"
    )

    lines.append(
        "=" * 80
    )

    lines.append(
        f"Số lượng kênh hiện tại: {current_count}"
    )

    if previous:
        sign = "+" if difference > 0 else ""

        lines.append(
            f"Chênh lệch so với lần chạy trước: "
            f"{sign}{difference}"
        )

    else:

        lines.append(
            "Chênh lệch so với lần chạy trước: "
            "Chưa có dữ liệu"
        )

    # --------------------------------------------------------
    # Kênh mới
    # --------------------------------------------------------

    lines.append("")
    lines.append(
        f"KÊNH MỚI ({len(new_channels)}):"
    )

    if new_channels:

        for item in new_channels:

            lines.append(
                f"  + {item['id']} | {item['name']}"
            )

    else:

        lines.append("  Không có.")

    # --------------------------------------------------------
    # Kênh mất
    # --------------------------------------------------------

    lines.append("")
    lines.append(
        f"KÊNH KHÔNG CÒN ({len(removed_channels)}):"
    )

    if removed_channels:

        for item in removed_channels:

            lines.append(
                f"  - {item['id']} | {item['name']}"
            )

    else:

        lines.append("  Không có.")

    # --------------------------------------------------------
    # EPG
    # --------------------------------------------------------

    lines.append("")
    lines.append(
        f"Số kênh có EPG: "
        f"{len(channels_with_epg)}"
    )

    lines.append(
        f"Số kênh không có EPG: "
        f"{len(channels_without_epg)}"
    )

    # --------------------------------------------------------
    # Chi tiết EPG
    # --------------------------------------------------------

    lines.append("")
    lines.append(
        "CHI TIẾT KÊNH CÓ EPG:"
    )

    for item in channels_with_epg:

        lines.append(
            f"  + {item['id']} | "
            f"{item['name']} | "
            f"{item['programmes']} chương trình"
        )

    # --------------------------------------------------------
    # Không có EPG
    # --------------------------------------------------------

    lines.append("")
    lines.append(
        "CHI TIẾT KÊNH KHÔNG CÓ EPG:"
    )

    if channels_without_epg:

        for item in channels_without_epg:

            lines.append(
                f"  - {item['id']} | "
                f"{item['name']} | "
                f"{item['reason']}"
            )

    else:

        lines.append("  Không có.")

    return (
        "\n".join(lines),
        json_data
    )


def save_log(log_text, json_data):
    """
    Lưu tối đa 7 lần chạy.

    Lần mới nhất nằm trên cùng.
    """

    previous_runs = load_previous_log()

    runs = [
        json_data
    ] + previous_runs

    runs = runs[:MAX_LOG_RUNS]

    # --------------------------------------------------------
    # Chúng ta lưu cả text + JSON.
    # --------------------------------------------------------

    # Text của run hiện tại đã được tạo ở build_log.
    # Các run cũ sẽ được phục dựng tối giản từ JSON.

    blocks = []

    # Current
    blocks.append(
        log_text
        + "\n\n"
        + "RUN_JSON_START\n"
        + json.dumps(
            json_data,
            ensure_ascii=False,
            indent=2
        )
        + "\nRUN_JSON_END"
    )

    # Old runs
    for old in runs[1:]:

        timestamp = old.get(
            "timestamp",
            ""
        )

        count = old.get(
            "channel_count",
            0
        )

        epg_count = old.get(
            "epg_channel_count",
            0
        )

        no_epg = old.get(
            "no_epg_channel_count",
            0
        )

        old_text = "\n".join([
            "=" * 80,
            f"TV360 - {timestamp}",
            "=" * 80,
            f"Số lượng kênh hiện tại: {count}",
            f"Số kênh có EPG: {epg_count}",
            f"Số kênh không có EPG: {no_epg}",
            "",
            "Dữ liệu chi tiết của lần chạy này được giữ trong JSON.",
            "",
            "RUN_JSON_START",
            json.dumps(
                old,
                ensure_ascii=False,
                indent=2
            ),
            "RUN_JSON_END",
        ])

        blocks.append(old_text)

    LOG_FILE.write_text(
        "\n\n".join(blocks),
        encoding="utf-8"
    )

    print(
        f"Đã cập nhật log: {LOG_FILE}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    start = time.time()

    print("=" * 70)
    print("TV360 EPG")
    print("=" * 70)

    print(
        f"Excel: {EXCEL_FILE}"
    )

    # --------------------------------------------------------
    # 1. Kiểm tra Excel
    # --------------------------------------------------------

    ensure_excel_file()

    # --------------------------------------------------------
    # 2. Lấy danh sách TV360
    # --------------------------------------------------------

    channels = get_channels_from_tv360()

    # --------------------------------------------------------
    # 3. Mở Excel
    # --------------------------------------------------------

    print(
        "Đang mở tv360channels.xlsx..."
    )

    wb = load_workbook(
        EXCEL_FILE,
        data_only=False
    )

    # --------------------------------------------------------
    # 4. Cập nhật Data A:E
    # --------------------------------------------------------

    update_data_sheet(
        wb,
        channels
    )

    # --------------------------------------------------------
    # 5. Đọc mapping từ Tham chiếu
    # --------------------------------------------------------

    reference_mapping = load_reference_mapping(
        wb
    )

    # --------------------------------------------------------
    # 6. Lưu Excel
    # --------------------------------------------------------

    wb.save(EXCEL_FILE)

    print(
        "Đã lưu tv360channels.xlsx."
    )

    # --------------------------------------------------------
    # 7. Lấy EPG
    # --------------------------------------------------------

    (
        programmes,
        channels_with_epg,
        channels_without_epg
    ) = collect_epg(
        channels,
        reference_mapping
    )

    # --------------------------------------------------------
    # 8. Tạo XML
    # --------------------------------------------------------

    programme_count = create_epg_xml(
        channels,
        reference_mapping,
        programmes
    )

    # --------------------------------------------------------
    # 9. Log
    # --------------------------------------------------------

    log_text, json_data = build_log(
        channels,
        channels_with_epg,
        channels_without_epg
    )

    json_data["programme_count"] = programme_count

    save_log(
        log_text,
        json_data
    )

    # --------------------------------------------------------
    # 10. Summary
    # --------------------------------------------------------

    elapsed = time.time() - start

    print("")
    print("=" * 70)
    print("HOÀN TẤT")
    print("=" * 70)

    print(
        f"Số kênh: {len(channels)}"
    )

    print(
        f"Kênh có EPG: "
        f"{len(channels_with_epg)}"
    )

    print(
        f"Kênh không có EPG: "
        f"{len(channels_without_epg)}"
    )

    print(
        f"Số chương trình: {programme_count}"
    )

    print(
        f"Thời gian chạy: {elapsed:.2f} giây"
    )

    print(
        f"XML: {EPG_FILE}"
    )

    print(
        f"Log: {LOG_FILE}"
    )


if __name__ == "__main__":
    main()
