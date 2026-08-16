import re
import time
import uuid
from datetime import datetime, timedelta

import pandas as pd
import pytz
import requests


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = "channel_list.xlsx"
OUTPUT_FILE = "epg.xml"

# EPG số ngày:
# 3 = hôm nay + ngày mai + ngày kia
EPG_DAYS = 3

# Múi giờ Việt Nam
TIMEZONE = pytz.timezone("Asia/Ho_Chi_Minh")

# Timeout mỗi request API
REQUEST_TIMEOUT = 30

# Số lần thử lại khi API lỗi
MAX_RETRIES = 3

# Thời gian chờ giữa các request
REQUEST_DELAY = 0.2

# User-Agent
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
}


# ============================================================
# HTTP SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# REQUEST JSON
# ============================================================

def request_json(url):
    """
    Gọi API và trả về JSON.

    Có retry khi:
    - timeout
    - connection error
    - HTTP error
    - response không phải JSON
    """

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            print(
                f"GET {url} "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )

            response = session.get(
                url,
                timeout=REQUEST_TIMEOUT
            )

            response.raise_for_status()

            # Parse JSON
            data = response.json()

            # Delay nhẹ giữa các request
            if REQUEST_DELAY > 0:
                time.sleep(REQUEST_DELAY)

            return data

        except requests.exceptions.Timeout as e:

            last_error = e

            print(
                f"[TIMEOUT] "
                f"attempt={attempt}/{MAX_RETRIES}"
            )

        except requests.exceptions.ConnectionError as e:

            last_error = e

            print(
                f"[CONNECTION ERROR] "
                f"attempt={attempt}/{MAX_RETRIES}"
            )

        except requests.exceptions.HTTPError as e:

            last_error = e

            print(
                f"[HTTP ERROR] "
                f"status={getattr(response, 'status_code', 'unknown')} "
                f"attempt={attempt}/{MAX_RETRIES}"
            )

            # In response nếu có
            try:
                print(
                    f"Response: "
                    f"{response.text[:500]}"
                )
            except Exception:
                pass

        except ValueError as e:

            # JSON decode error
            last_error = e

            print(
                f"[JSON ERROR] "
                f"attempt={attempt}/{MAX_RETRIES}"
            )

            try:
                print(
                    f"Response: "
                    f"{response.text[:500]}"
                )
            except Exception:
                pass

        except Exception as e:

            last_error = e

            print(
                f"[REQUEST ERROR] "
                f"{type(e).__name__}: {e}"
            )

        # Retry
        if attempt < MAX_RETRIES:

            wait_time = attempt * 2

            print(
                f"Retry sau {wait_time} giây..."
            )

            time.sleep(wait_time)

    print()
    print(
        f"[API FAILED] "
        f"Không thể lấy dữ liệu sau "
        f"{MAX_RETRIES} lần thử."
    )

    print(
        f"URL: {url}"
    )

    print(
        f"Lỗi cuối: {last_error}"
    )

    return None


# ============================================================
# XML ESCAPE
# ============================================================

def xml_escape(value):
    """
    Escape XML đúng một lần.
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


# ============================================================
# FORMAT TITLE
# ============================================================

def format_title_string(title_str):
    """
    Chuẩn hóa tên chương trình.
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

    # Chuẩn hóa khoảng trắng
    title_str = re.sub(
        r"\s+",
        " ",
        title_str
    )

    # Dấu phẩy
    title_str = re.sub(
        r"\s*,\s*",
        ", ",
        title_str
    )

    # Dấu :
    title_str = re.sub(
        r"\s*:\s*",
        ": ",
        title_str
    )

    # Dấu -
    title_str = re.sub(
        r"\s*-\s*",
        " - ",
        title_str
    )

    return title_str.strip()


# ============================================================
# FORMAT XMLTV DATETIME
# ============================================================

def format_xml_datetime(dt):
    """
    XMLTV datetime:

    YYYYMMDDHHMMSS +0700
    """

    return dt.strftime(
        "%Y%m%d%H%M%S %z"
    )


# ============================================================
# GET CHANNEL LIST FROM MYTV API
# ============================================================

def get_channel_list_from_api(api_uuid):

    url = (
        "https://apigw.mytv.vn/api/v1/channel"
        f"?cate_id=undefined&uuid={api_uuid}"
    )

    data = request_json(url)

    if data is None:
        print(
            "[ERROR] Không nhận được dữ liệu "
            "channel từ MyTV API."
        )

        return {}

    if not isinstance(data, dict):

        print(
            "[ERROR] Response channel API "
            "không phải dictionary."
        )

        print(
            f"Response type: {type(data)}"
        )

        return {}

    response_data = data.get("data")

    if response_data is None:

        print(
            "[ERROR] channel API trả data=None."
        )

        return {}

    # Một số API có thể trả list trực tiếp
    if isinstance(response_data, list):

        channel_list = response_data

    # Một số trường hợp data là dict
    elif isinstance(response_data, dict):

        channel_list = (
            response_data.get("channels")
            or response_data.get("list")
            or []
        )

    else:

        print(
            "[ERROR] channel API data "
            "không đúng định dạng."
        )

        print(
            f"data type: {type(response_data)}"
        )

        return {}

    api_channels = {}

    for channel in channel_list:

        if not isinstance(channel, dict):
            continue

        channel_id = channel.get(
            "channel_id"
        )

        channel_name = channel.get(
            "name"
        )

        if channel_id is None:
            continue

        if channel_name is None:
            continue

        channel_id = str(
            channel_id
        ).strip()

        channel_name = str(
            channel_name
        ).strip()

        if not channel_id:
            continue

        if not channel_name:
            continue

        api_channels[channel_id] = channel_name

    print()
    print(
        "========================================"
    )

    print(
        f"MyTV API trả về "
        f"{len(api_channels)} kênh"
    )

    print(
        "========================================"
    )

    return api_channels


# ============================================================
# LOAD EXCEL CHANNELS
# ============================================================

def load_excel_channels():

    try:

        df = pd.read_excel(
            EXCEL_FILE
        )

    except FileNotFoundError:

        raise FileNotFoundError(
            f"Không tìm thấy file "
            f"{EXCEL_FILE}. "
            f"Hãy đặt file này ở thư mục gốc "
            f"repository."
        )

    except Exception as e:

        raise RuntimeError(
            f"Không thể đọc {EXCEL_FILE}: {e}"
        )

    print()
    print(
        "========================================"
    )

    print(
        f"Đọc file: {EXCEL_FILE}"
    )

    print(
        "========================================"
    )

    print(
        f"Số dòng Excel: {len(df)}"
    )

    print(
        f"Các cột: {list(df.columns)}"
    )

    print()

    # --------------------------------------------------------
    # Kiểm tra cột
    # --------------------------------------------------------

    required_columns = [
        "channel",
        "name"
    ]

    for column in required_columns:

        if column not in df.columns:

            raise RuntimeError(
                f"channel_list.xlsx thiếu "
                f"cột bắt buộc: {column}"
            )

    channels = []

    for _, row in df.iterrows():

        # ----------------------------------------------------
        # Channel ID
        # ----------------------------------------------------

        value = row.get(
            "channel"
        )

        if pd.isna(value):
            continue

        excel_channel_id = str(
            value
        ).strip()

        if not excel_channel_id:
            continue

        # ----------------------------------------------------
        # Name
        # ----------------------------------------------------

        value = row.get(
            "name"
        )

        if pd.isna(value):
            continue

        name = str(
            value
        ).strip()

        if not name:
            continue

        # ----------------------------------------------------
        # Display name
        # ----------------------------------------------------

        display_name = ""

        if "display-name" in df.columns:

            value = row.get(
                "display-name"
            )

            if not pd.isna(value):

                display_name = str(
                    value
                ).strip()

        # Nếu display-name trống
        # thì dùng name
        if not display_name:
            display_name = name

        # ----------------------------------------------------
        # Display number
        # ----------------------------------------------------

        display_number = ""

        if "display-number" in df.columns:

            value = row.get(
                "display-number"
            )

            if not pd.isna(value):

                display_number = str(
                    value
                ).strip()

        # ----------------------------------------------------
        # Add channel
        # ----------------------------------------------------

        channels.append({

            "excel_channel_id":
                excel_channel_id,

            "name":
                name,

            "display-name":
                display_name,

            "display-number":
                display_number
        })

    print(
        f"Đã đọc {len(channels)} kênh "
        f"từ Excel."
    )

    return channels


# ============================================================
# MATCH CHANNELS
# ============================================================

def match_channels(
    excel_channels,
    api_channels
):

    matched = []

    print()
    print(
        "========================================"
    )

    print(
        "Đối chiếu kênh Excel với MyTV API"
    )

    print(
        "========================================"
    )

    for channel in excel_channels:

        excel_name = channel[
            "name"
        ].strip()

        found_api_id = None

        # ----------------------------------------------------
        # Match chính xác theo tên
        # ----------------------------------------------------

        for api_channel_id, api_name in api_channels.items():

            if (
                api_name.strip().lower()
                ==
                excel_name.lower()
            ):

                found_api_id = (
                    api_channel_id
                )

                break

        # ----------------------------------------------------
        # Không tìm thấy
        # ----------------------------------------------------

        if found_api_id is None:

            print(
                f"[KHÔNG KHỚP] "
                f"{channel['excel_channel_id']} "
                f"- {excel_name}"
            )

            continue

        item = dict(channel)

        item[
            "api_channel_id"
        ] = found_api_id

        matched.append(item)

        print(
            f"[OK] "
            f"{channel['excel_channel_id']} "
            f"- {excel_name} "
            f"=> API {found_api_id}"
        )

    print()
    print(
        "========================================"
    )

    print(
        f"Khớp được "
        f"{len(matched)}/"
        f"{len(excel_channels)} kênh"
    )

    print(
        "========================================"
    )

    return matched


# ============================================================
# GET SCHEDULE FOR CHANNEL
# ============================================================

def get_schedule_for_channel(
    api_channel_id,
    date_obj,
    api_uuid
):

    date_str = date_obj.strftime(
        "%Y-%m-%d"
    )

    url = (
        "https://apigw.mytv.vn/api/v1/channel/"
        f"{api_channel_id}/schedule"
        f"?date={date_str}"
        f"&uuid={api_uuid}"
    )

    data = request_json(url)

    # ========================================================
    # API ERROR
    # ========================================================

    if data is None:

        print(
            f"[API ERROR] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        return []

    # ========================================================
    # RESPONSE MUST BE DICT
    # ========================================================

    if not isinstance(data, dict):

        print(
            f"[INVALID API RESPONSE] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        print(
            f"Response type: "
            f"{type(data)}"
        )

        print(
            f"Response: {data}"
        )

        return []

    # ========================================================
    # GET data
    # ========================================================

    response_data = data.get(
        "data"
    )

    # --------------------------------------------------------
    # data = None
    # --------------------------------------------------------

    if response_data is None:

        print(
            f"[NO SCHEDULE DATA] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        return []

    # --------------------------------------------------------
    # data must be dict
    # --------------------------------------------------------

    if not isinstance(
        response_data,
        dict
    ):

        print(
            f"[INVALID DATA] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        print(
            f"data type: "
            f"{type(response_data)}"
        )

        print(
            f"data: "
            f"{response_data}"
        )

        return []

    # ========================================================
    # GET schedule
    # ========================================================

    schedule = response_data.get(
        "schedule"
    )

    # --------------------------------------------------------
    # schedule = None
    # --------------------------------------------------------

    if schedule is None:

        print(
            f"[NO SCHEDULE] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        return []

    # --------------------------------------------------------
    # schedule must be list
    # --------------------------------------------------------

    if not isinstance(
        schedule,
        list
    ):

        print(
            f"[INVALID SCHEDULE] "
            f"channel={api_channel_id}, "
            f"date={date_str}"
        )

        print(
            f"schedule type: "
            f"{type(schedule)}"
        )

        print(
            f"schedule: "
            f"{schedule}"
        )

        return []

    # ========================================================
    # PROCESS PROGRAMMES
    # ========================================================

    result = []

    for item in schedule:

        # ----------------------------------------------------
        # Programme must be dict
        # ----------------------------------------------------

        if not isinstance(
            item,
            dict
        ):
            continue

        # ----------------------------------------------------
        # Title
        # ----------------------------------------------------

        raw_title = item.get(
            "title",
            ""
        )

        # ----------------------------------------------------
        # Time
        # ----------------------------------------------------

        time_str = item.get(
            "time",
            ""
        )

        if not time_str:

            continue

        title = format_title_string(
            raw_title
        )

        # ====================================================
        # PARSE DATETIME
        # ====================================================

        try:

            naive_datetime = (
                datetime.strptime(
                    f"{date_str} {time_str}",
                    "%Y-%m-%d %H:%M"
                )
            )

            local_datetime = (
                TIMEZONE.localize(
                    naive_datetime
                )
            )

        except ValueError as e:

            print(
                f"[INVALID TIME] "
                f"channel={api_channel_id}, "
                f"date={date_str}, "
                f"time={time_str}, "
                f"error={e}"
            )

            continue

        result.append({

            "title":
                title,

            "local_start_time":
                local_datetime
        })

    # ========================================================
    # SORT
    # ========================================================

    result.sort(
        key=lambda x:
        x["local_start_time"]
    )

    print(
        f"[EPG OK] "
        f"channel={api_channel_id}, "
        f"date={date_str}, "
        f"programmes={len(result)}"
    )

    return result


# ============================================================
# BUILD PROGRAMMES
# ============================================================

def build_programmes(
    channel,
    date_obj,
    api_uuid
):

    api_channel_id = channel[
        "api_channel_id"
    ]

    schedule = (
        get_schedule_for_channel(
            api_channel_id,
            date_obj,
            api_uuid
        )
    )

    if not schedule:

        print(
            f"[NO EPG] "
            f"{channel['name']} "
            f"{date_obj}"
        )

        return []

    programmes = []

    for index, program in enumerate(
        schedule
    ):

        start_dt = program[
            "local_start_time"
        ]

        # ====================================================
        # END TIME
        # ====================================================

        if index + 1 < len(
            schedule
        ):

            next_start_dt = (
                schedule[
                    index + 1
                ][
                    "local_start_time"
                ]
            )

            duration_seconds = (
                next_start_dt -
                start_dt
            ).total_seconds()

            duration_minutes = int(
                duration_seconds / 60
            )

            # ------------------------------------------------
            # Tránh duration <= 0
            # ------------------------------------------------

            if duration_minutes <= 0:

                duration_minutes = 30

            # ------------------------------------------------
            # Giới hạn thời lượng bất thường
            # ------------------------------------------------

            if duration_minutes > 24 * 60:

                duration_minutes = 30

            end_dt = (
                next_start_dt
            )

        else:

            # Chương trình cuối cùng
            # mặc định 30 phút

            duration_minutes = 30

            end_dt = (
                start_dt +
                timedelta(
                    minutes=duration_minutes
                )
            )

        # ====================================================
        # ADD
        # ====================================================

        programmes.append({

            "start":
                format_xml_datetime(
                    start_dt
                ),

            "stop":
                format_xml_datetime(
                    end_dt
                ),

            "title":
                program["title"],

            "length":
                duration_minutes
        })

    return programmes


# ============================================================
# CREATE EPG XML
# ============================================================

def create_epg_xml(
    matched_channels,
    output_file,
    api_uuid
):

    now = datetime.now(
        TIMEZONE
    )

    first_date = now.date()

    print()
    print(
        "========================================"
    )

    print(
        f"Tạo EPG {EPG_DAYS} ngày"
    )

    print(
        f"Từ: {first_date}"
    )

    print(
        f"Đến: "
        f"{first_date + timedelta(days=EPG_DAYS - 1)}"
    )

    print(
        "========================================"
    )

    # ========================================================
    # XML
    # ========================================================

    xml_lines = []

    xml_lines.append(
        '<?xml version="1.0" encoding="UTF-8"?>'
    )

    xml_lines.append(
        f'<tv '
        f'date="{now.strftime("%d-%m-%Y")}" '
        f'source-info-name="Ngan Phuc" '
        f'generator-info-name="MyTV EPG GitHub"'
        f'>'
    )

    # ========================================================
    # CHANNEL
    # ========================================================

    for channel in matched_channels:

        channel_id = xml_escape(
            channel[
                "excel_channel_id"
            ]
        )

        display_name = xml_escape(
            channel.get(
                "display-name",
                ""
            )
        )

        display_number = xml_escape(
            channel.get(
                "display-number",
                ""
            )
        )

        # ----------------------------------------------------
        # Channel
        # ----------------------------------------------------

        xml_lines.append(
            f'  <channel '
            f'id="{channel_id}">'
        )

        # ----------------------------------------------------
        # Display name
        # ----------------------------------------------------

        xml_lines.append(
            f'    <display-name '
            f'lang="vi">'
            f'{display_name}'
            f'</display-name>'
        )

        # ----------------------------------------------------
        # Display number
        # ----------------------------------------------------

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
    # PROGRAMMES
    # ========================================================

    total_programmes = 0

    total_no_epg = 0

    for day_offset in range(
        EPG_DAYS
    ):

        date_obj = (
            first_date +
            timedelta(
                days=day_offset
            )
        )

        print()
        print(
            "========================================"
        )

        print(
            f"NGÀY "
            f"{day_offset + 1}/"
            f"{EPG_DAYS}: "
            f"{date_obj}"
        )

        print(
            "========================================"
        )

        for channel in matched_channels:

            try:

                programmes = (
                    build_programmes(
                        channel,
                        date_obj,
                        api_uuid
                    )
                )

            except Exception as e:

                # ------------------------------------------------
                # Một kênh lỗi không làm chết cả workflow
                # ------------------------------------------------

                print(
                    f"[CHANNEL ERROR] "
                    f"{channel['name']} "
                    f"({channel['api_channel_id']}) "
                    f"- {date_obj}"
                )

                print(
                    f"{type(e).__name__}: {e}"
                )

                total_no_epg += 1

                continue

            # ------------------------------------------------
            # Không có programme
            # ------------------------------------------------

            if not programmes:

                total_no_epg += 1

                continue

            channel_id = xml_escape(
                channel[
                    "excel_channel_id"
                ]
            )

            # =================================================
            # WRITE PROGRAMMES
            # =================================================

            for programme in programmes:

                title = xml_escape(
                    programme[
                        "title"
                    ]
                )

                length = programme[
                    "length"
                ]

                start = programme[
                    "start"
                ]

                stop = programme[
                    "stop"
                ]

                # ------------------------------------------------
                # programme
                # ------------------------------------------------

                xml_lines.append(
                    f'  <programme '
                    f'start="{start}" '
                    f'stop="{stop}" '
                    f'channel="{channel_id}">'
                )

                # ------------------------------------------------
                # title
                # ------------------------------------------------

                xml_lines.append(
                    f'    <title '
                    f'lang="vi">'
                    f'{title}'
                    f'</title>'
                )

                # ------------------------------------------------
                # length
                # ------------------------------------------------

                xml_lines.append(
                    f'    <length '
                    f'lang="vi">'
                    f'{length}'
                    f'</length>'
                )

                xml_lines.append(
                    '  </programme>'
                )

                total_programmes += 1

    # ========================================================
    # CLOSE XML
    # ========================================================

    xml_lines.append(
        '</tv>'
    )

    xml_content = "\n".join(
        xml_lines
    )

    # ========================================================
    # WRITE FILE
    # ========================================================

    try:

        with open(
            output_file,
            "w",
            encoding="utf-8",
            newline="\n"
        ) as file:

            file.write(
                xml_content
            )

    except Exception as e:

        raise RuntimeError(
            f"Không thể ghi "
            f"{output_file}: {e}"
        )

    # ========================================================
    # STATISTICS
    # ========================================================

    print()
    print(
        "========================================"
    )

    print(
        "           HOÀN THÀNH EPG"
    )

    print(
        "========================================"
    )

    print(
        f"Số kênh:       "
        f"{len(matched_channels)}"
    )

    print(
        f"Số ngày:       "
        f"{EPG_DAYS}"
    )

    print(
        f"Programme:     "
        f"{total_programmes}"
    )

    print(
        f"Kênh/ngày lỗi: "
        f"{total_no_epg}"
    )

    print(
        f"Output:        "
        f"{output_file}"
    )

    print(
        "========================================"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print(
        "========================================"
    )

    print(
        "       MyTV EPG Generator"
    )

    print(
        "       GitHub Actions"
    )

    print(
        "       EPG 3 NGÀY"
    )

    print(
        "========================================"
    )

    # ========================================================
    # UUID
    # ========================================================

    api_uuid = str(
        uuid.uuid4()
    )

    print(
        f"UUID: {api_uuid}"
    )

    print(
        f"Timezone: "
        f"{TIMEZONE}"
    )

    print(
        f"EPG days: "
        f"{EPG_DAYS}"
    )

    print()

    # ========================================================
    # 1. LOAD EXCEL
    # ========================================================

    excel_channels = (
        load_excel_channels()
    )

    if not excel_channels:

        raise RuntimeError(
            "Không có kênh nào trong "
            "channel_list.xlsx"
        )

    # ========================================================
    # 2. GET MYTV CHANNELS
    # ========================================================

    api_channels = (
        get_channel_list_from_api(
            api_uuid
        )
    )

    if not api_channels:

        raise RuntimeError(
            "Không lấy được danh sách "
            "kênh từ MyTV API."
        )

    # ========================================================
    # 3. MATCH
    # ========================================================

    matched_channels = (
        match_channels(
            excel_channels,
            api_channels
        )
    )

    if not matched_channels:

        raise RuntimeError(
            "Không có kênh nào khớp "
            "giữa channel_list.xlsx "
            "và MyTV API."
        )

    # ========================================================
    # 4. CREATE EPG
    # ========================================================

    create_epg_xml(
        matched_channels,
        OUTPUT_FILE,
        api_uuid
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()
