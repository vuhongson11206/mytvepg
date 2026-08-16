import json
import re
import uuid
from datetime import datetime, timedelta

import pandas as pd
import pytz
import requests


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = "."
EXCEL_FILE = "channel_list.xlsx"
OUTPUT_FILE = "epg.xml"

TIMEZONE = pytz.timezone("Asia/Ho_Chi_Minh")

# Số ngày EPG muốn lấy, tính từ hôm nay
EPG_DAYS = 3

# Timeout cho mỗi request API
REQUEST_TIMEOUT = 30

# Số lần retry khi API lỗi
MAX_RETRIES = 3

# User-Agent
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    )
}


# ============================================================
# SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# HELPERS
# ============================================================

def request_json(url):
    """
    Gọi API và trả về JSON.
    Có retry nếu request thất bại.
    """

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"GET {url} (attempt {attempt}/{MAX_RETRIES})")

            response = session.get(
                url,
                timeout=REQUEST_TIMEOUT
            )

            response.raise_for_status()

            return response.json()

        except Exception as e:
            last_error = e
            print(f"Request lỗi: {e}")

            if attempt < MAX_RETRIES:
                import time
                time.sleep(2 * attempt)

    print(f"API thất bại sau {MAX_RETRIES} lần: {url}")
    print(f"Lỗi cuối: {last_error}")

    return None


def xml_escape(value):
    """
    Escape dữ liệu trước khi đưa vào XML.
    """

    if value is None:
        return ""

    value = str(value)

    value = value.replace("&", "&amp;")
    value = value.replace("<", "&lt;")
    value = value.replace(">", "&gt;")
    value = value.replace('"', "&quot;")
    value = value.replace("'", "&apos;")

    return value


def format_title_string(title_str):
    """
    Chuẩn hóa tiêu đề chương trình.
    """

    if title_str is None:
        return ""

    title_str = str(title_str)

    title_str = (
        title_str
        .replace("\r", "")
        .replace("\n", " ")
        .replace("\t", " ")
        .strip()
    )

    # Dấu phẩy
    title_str = re.sub(r"\s*,\s*", ", ", title_str)

    # Dấu :
    title_str = re.sub(r"\s*:\s*", ": ", title_str)

    # Dấu -
    title_str = re.sub(r"\s*-\s*", " - ", title_str)

    # Khoảng trắng kép
    title_str = re.sub(r"\s+", " ", title_str).strip()

    return title_str


def format_xml_datetime(dt):
    """
    XMLTV datetime:
    YYYYMMDDHHMMSS +0700
    """

    return dt.strftime("%Y%m%d%H%M%S %z")


# ============================================================
# GET CHANNEL LIST FROM MYTV API
# ============================================================

def get_channel_list_from_api(api_uuid):
    url = (
        "https://apigw.mytv.vn/api/v1/channel"
        f"?cate_id=undefined&uuid={api_uuid}"
    )

    data = request_json(url)

    if not data:
        return {}

    api_channels = {}

    if isinstance(data.get("data"), list):
        for channel in data["data"]:

            channel_id = channel.get("channel_id")
            channel_name = channel.get("name")

            if channel_id is None or channel_name is None:
                continue

            api_channels[str(channel_id)] = str(channel_name).strip()

    print()
    print("========================================")
    print(f"API trả về {len(api_channels)} kênh")
    print("========================================")

    return api_channels


# ============================================================
# READ CHANNEL LIST FROM EXCEL
# ============================================================

def load_excel_channels():
    try:
        df = pd.read_excel(EXCEL_FILE)

    except FileNotFoundError:
        raise FileNotFoundError(
            f"Không tìm thấy {EXCEL_FILE}. "
            "Hãy đặt channel_list.xlsx ở thư mục gốc repository."
        )

    except Exception as e:
        raise RuntimeError(
            f"Không thể đọc {EXCEL_FILE}: {e}"
        )

    required_columns = ["channel", "name"]

    for column in required_columns:
        if column not in df.columns:
            raise RuntimeError(
                f"channel_list.xlsx thiếu cột bắt buộc: {column}"
            )

    channels = []

    for _, row in df.iterrows():

        excel_channel_id = str(row["channel"]).strip()

        if not excel_channel_id or excel_channel_id.lower() == "nan":
            continue

        name = str(row["name"]).strip()

        if not name or name.lower() == "nan":
            continue

        display_name = ""

        if "display-name" in df.columns:
            value = row["display-name"]

            if pd.notna(value):
                display_name = str(value).strip()

        display_number = ""

        if "display-number" in df.columns:
            value = row["display-number"]

            if pd.notna(value):
                display_number = str(value).strip()

        channels.append({
            "excel_channel_id": excel_channel_id,
            "name": name,
            "display-name": display_name,
            "display-number": display_number
        })

    print()
    print("========================================")
    print(f"Excel có {len(channels)} kênh")
    print("========================================")

    return channels


# ============================================================
# MATCH EXCEL CHANNEL WITH MYTV API CHANNEL
# ============================================================

def match_channels(excel_channels, api_channels):
    matched = []

    # API:
    # {
    #     "channel_id": "123",
    #     "name": "VTV1"
    # }

    for excel_channel in excel_channels:

        excel_name = excel_channel["name"].strip()

        found_api_id = None

        for api_channel_id, api_name in api_channels.items():

            if api_name.strip() == excel_name:
                found_api_id = api_channel_id
                break

        if found_api_id is None:

            print(
                f"[KHÔNG KHỚP] "
                f"{excel_channel['excel_channel_id']} - "
                f"{excel_name}"
            )

            continue

        item = dict(excel_channel)

        item["api_channel_id"] = found_api_id

        matched.append(item)

        print(
            f"[OK] "
            f"{excel_channel['excel_channel_id']} - "
            f"{excel_name} "
            f"=> API {found_api_id}"
        )

    print()
    print("========================================")
    print(f"Khớp được {len(matched)}/{len(excel_channels)} kênh")
    print("========================================")

    return matched


# ============================================================
# GET SCHEDULE
# ============================================================

def get_schedule_for_channel(
    api_channel_id,
    date_obj,
    api_uuid
):
    date_str = date_obj.strftime("%Y-%m-%d")

    url = (
        f"https://apigw.mytv.vn/api/v1/channel/"
        f"{api_channel_id}/schedule"
        f"?date={date_str}&uuid={api_uuid}"
    )

    data = request_json(url)

    if not data:
        return []

    result = []

    schedule = (
        data.get("data", {})
        .get("schedule", [])
    )

    if not isinstance(schedule, list):
        return []

    for item in schedule:

        raw_title = item.get("title", "")
        time_str = item.get("time", "")

        if not time_str:
            continue

        title = format_title_string(raw_title)

        try:

            naive_datetime = datetime.strptime(
                f"{date_str} {time_str}",
                "%Y-%m-%d %H:%M"
            )

            local_datetime = TIMEZONE.localize(
                naive_datetime
            )

        except ValueError:
            print(
                f"Không thể parse thời gian: "
                f"{date_str} {time_str}"
            )
            continue

        result.append({
            "title": title,
            "local_start_time": local_datetime
        })

    # Sắp xếp theo giờ
    result.sort(
        key=lambda x: x["local_start_time"]
    )

    return result


# ============================================================
# CREATE PROGRAMMES
# ============================================================

def build_programmes(
    channel,
    date_obj,
    api_uuid
):
    api_channel_id = channel["api_channel_id"]

    schedule = get_schedule_for_channel(
        api_channel_id,
        date_obj,
        api_uuid
    )

    if not schedule:
        print(
            f"[NO EPG] "
            f"{channel['name']} "
            f"{date_obj}"
        )
        return []

    programmes = []

    for index, program in enumerate(schedule):

        start_dt = program["local_start_time"]

        # ----------------------------------------------------
        # End time
        # ----------------------------------------------------

        if index + 1 < len(schedule):

            next_start_dt = (
                schedule[index + 1]["local_start_time"]
            )

            duration = (
                next_start_dt - start_dt
            ).total_seconds() / 60

            duration_minutes = int(duration)

            # Tránh thời lượng <= 0
            if duration_minutes <= 0:
                duration_minutes = 30

            end_dt = next_start_dt

        else:

            # Chương trình cuối cùng:
            # mặc định 30 phút
            duration_minutes = 30

            end_dt = (
                start_dt +
                timedelta(minutes=duration_minutes)
            )

        programmes.append({
            "start": format_xml_datetime(start_dt),
            "stop": format_xml_datetime(end_dt),
            "title": program["title"],
            "length": duration_minutes
        })

    return programmes


# ============================================================
# CREATE XML
# ============================================================

def create_epg_xml(
    matched_channels,
    output_file,
    api_uuid
):
    now = datetime.now(TIMEZONE)

    first_date = now.date()

    print()
    print("========================================")
    print(
        f"Tạo EPG {EPG_DAYS} ngày "
        f"từ {first_date}"
    )
    print("========================================")

    xml_lines = []

    xml_lines.append(
        '<?xml version="1.0" encoding="UTF-8"?>'
    )

    xml_lines.append(
        f'<tv '
        f'date="{now.strftime("%d-%m-%Y")}" '
        f'source-info-name="Ngan Phuc" '
        f'generator-info-name="MyTV EPG - '
        f'Cap nhat luc {now.strftime("%H:%M:%S - %d/%m/%Y")}"'
        f'>'
    )

    # ========================================================
    # CHANNEL
    # ========================================================

    for channel in matched_channels:

        channel_id = xml_escape(
            channel["excel_channel_id"]
        )

        display_name = xml_escape(
            channel.get("display-name", "")
        )

        display_number = xml_escape(
            channel.get("display-number", "")
        )

        xml_lines.append(
            f'  <channel id="{channel_id}">'
        )

        xml_lines.append(
            f'    <display-name lang="vi">'
            f'{display_name}'
            f'</display-name>'
        )

        if display_number:
            xml_lines.append(
                f'    <display-number>'
                f'{display_number}'
                f'</display-number>'
            )

        xml_lines.append(
            '  </channel>'
        )

    # ========================================================
    # PROGRAMME
    # ========================================================

    total_programmes = 0

    for day_offset in range(EPG_DAYS):

        date_obj = first_date + timedelta(
            days=day_offset
        )

        print()
        print(
            f"===== NGÀY {day_offset + 1}/"
            f"{EPG_DAYS}: {date_obj} ====="
        )

        for channel in matched_channels:

            programmes = build_programmes(
                channel,
                date_obj,
                api_uuid
            )

            channel_id = xml_escape(
                channel["excel_channel_id"]
            )

            for programme in programmes:

                title = xml_escape(
                    programme["title"]
                )

                length = programme["length"]

                xml_lines.append(
                    f'  <programme '
                    f'start="{programme["start"]}" '
                    f'stop="{programme["stop"]}" '
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
                    f'{length} phút'
                    f'</length>'
                )

                xml_lines.append(
                    '  </programme>'
                )

                total_programmes += 1

    # ========================================================
    # END XML
    # ========================================================

    xml_lines.append("</tv>")

    xml_content = "\n".join(xml_lines)

    # Ghi file UTF-8
    with open(
        output_file,
        "w",
        encoding="utf-8",
        newline="\n"
    ) as file:

        file.write(xml_content)

    print()
    print("========================================")
    print("HOÀN THÀNH")
    print("========================================")
    print(f"Kênh:        {len(matched_channels)}")
    print(f"Số ngày:     {EPG_DAYS}")
    print(f"Programme:   {total_programmes}")
    print(f"Output:      {output_file}")
    print("========================================")


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("========================================")
    print("       MyTV EPG Generator")
    print("       EPG 3 ngày")
    print("========================================")

    # UUID mới cho mỗi lần chạy
    api_uuid = str(uuid.uuid4())

    # --------------------------------------------------------
    # 1. Đọc channel_list.xlsx
    # --------------------------------------------------------

    excel_channels = load_excel_channels()

    if not excel_channels:
        raise RuntimeError(
            "Không có kênh nào trong channel_list.xlsx"
        )

    # --------------------------------------------------------
    # 2. Lấy danh sách kênh từ MyTV
    # --------------------------------------------------------

    api_channels = get_channel_list_from_api(
        api_uuid
    )

    if not api_channels:
        raise RuntimeError(
            "Không lấy được danh sách kênh từ MyTV API"
        )

    # --------------------------------------------------------
    # 3. Match channel
    # --------------------------------------------------------

    matched_channels = match_channels(
        excel_channels,
        api_channels
    )

    if not matched_channels:
        raise RuntimeError(
            "Không có kênh nào khớp giữa Excel và MyTV API"
        )

    # --------------------------------------------------------
    # 4. Generate EPG
    # --------------------------------------------------------

    create_epg_xml(
        matched_channels,
        OUTPUT_FILE,
        api_uuid
    )


if __name__ == "__main__":
    main()