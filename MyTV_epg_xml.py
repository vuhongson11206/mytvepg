import requests
import json
import pandas as pd
import uuid
from datetime import datetime, date, timedelta, timezone
import pytz
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

def format_title_string(title_str):
    """
    Chỉnh sửa định dạng chuỗi tiêu đề theo các quy tắc đã cho.
    - Nếu trước dấu "," hoặc dấu ":" có khoảng trắng thì xóa khoảng trắng đó.
      Nếu sau dấu "," hoặc dấu ":" không có khoảng trắng thêm 1 khoảng trắng ngay sau dấu đó.
    - Nếu trước dấu "-" không có khoảng trắng thêm 1 khoảng trắng ngay trước dấu đó.
      Nếu sau dấu "-" không có khoảng trắng thêm 1 khoảng trắng ngay sau dấu đó.
    """
    if not isinstance(title_str, str): # Đảm bảo title_str là chuỗi
        return str(title_str)

    # Quy tắc cho dấu phẩy và dấu hai chấm
    # Loại bỏ khoảng trắng trước dấu ',' hoặc ':'
    title_str = title_str.replace(' ,', ',').replace(' :', ':')
    # Thêm khoảng trắng sau dấu ',' hoặc ':' nếu chưa có
    title_str = title_str.replace(',', ', ').replace(':', ': ')

    # Quy tắc cho dấu gạch nối
    # Thêm khoảng trắng trước và sau dấu '-' nếu chưa có
    # Sử dụng regex để xử lý các trường hợp phức tạp hơn như "text-text"
    import re
    # Thêm khoảng trắng trước '-' nếu chưa có
    title_str = re.sub(r'(\S)-', r'\1 -', title_str)
    # Thêm khoảng trắng sau '-' nếu chưa có
    title_str = re.sub(r'-(\S)', r'- \1', title_str)
    # Loại bỏ khoảng trắng thừa nếu có trường hợp ' - '
    title_str = title_str.replace(' - ', ' - ').replace(' -  ', ' - ')

    # Loại bỏ khoảng trắng kép thành một khoảng trắng đơn
    title_str = re.sub(r'\s+', ' ', title_str).strip()

    return title_str

def get_channel_list_from_api(uuid_str):
    url = f"https://apigw.mytv.vn/api/v1/channel?cate_id=undefined&uuid={uuid_str}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        api_channels = {}
        if 'data' in data and isinstance(data['data'], list):
            for channel in data['data']:
                api_channels[str(channel.get('channel_id'))] = channel.get('name')
        print("Dữ liệu kênh từ API:", api_channels)
        return api_channels
    except requests.exceptions.RequestException as e:
        print(f"Lỗi khi lấy danh sách kênh từ API: {e}")
        return {}
    except json.JSONDecodeError:
        print("Lỗi khi giải mã JSON danh sách kênh từ API.")
        return {}
        
def auto_adjust_column_width(excel_file, sheet_name):
    try:
        workbook = load_workbook(excel_file)
        sheet = workbook[sheet_name]
        for column_cells in sheet.columns:
            max_length = 0
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width
        workbook.save(excel_file)
        print(f"Đã tự động điều chỉnh kích thước cột trong file: {excel_file}, sheet: {sheet_name}")
    except Exception as e:
        print(f"Lỗi khi tự động điều chỉnh kích thước cột: {e}")        

def compare_channel_lists(excel_file, api_channels):
    try:
        df = pd.read_excel(excel_file)
        excel_channels_data = []
        if 'channel' in df.columns and 'name' in df.columns:
            for index, row in df.iterrows():
                channel_excel_id = str(row['channel']).strip()
                name = str(row['name']).strip()
                excel_channels_data.append({'channel': channel_excel_id, 'name': name})
        print("Dữ liệu kênh từ Excel:", excel_channels_data)

        matched_channels = {}
        for api_channel_id, api_name in api_channels.items():
            api_name_str = str(api_name).strip()
            for excel_channel in excel_channels_data:
                if api_name_str == excel_channel['name']:
                    matched_channels[excel_channel['channel']] = {
                        'name': excel_channel['name'],
                        'display-name': df.loc[df['channel'] == excel_channel['channel'], 'display-name'].iloc[0] if 'display-name' in df.columns and not df.loc[df['channel'] == excel_channel['channel'], 'display-name'].empty else None,
                        'display-number': df.loc[df['channel'] == excel_channel['channel'], 'display-number'].iloc[0] if 'display-number' in df.columns and not df.loc[df['channel'] == excel_channel['channel'], 'display-number'].empty else None,
                        'excel_channel_id': excel_channel['channel'],
                        'api_channel_id': api_channel_id # Lưu trữ ID từ API nếu cần
                    }
                    break # Tìm thấy tên khớp, không cần tìm tiếp cho tên API này
        print("Các kênh khớp nhau (dựa trên tên):", matched_channels)

        # Tự động điều chỉnh kích thước cột sau khi (có thể) ghi dữ liệu vào Excel
        # Lưu ý: Đoạn code hiện tại chỉ đọc Excel, nếu bạn có ghi lại Excel, hãy gọi hàm này sau khi ghi.
        auto_adjust_column_width(excel_file, df.sheet_name if hasattr(df, 'sheet_name') else 'Sheet1')
        
        return matched_channels

    except FileNotFoundError:
        print(f"Không tìm thấy file Excel: {excel_file}")
        return {}
    except Exception as e:
        print(f"Lỗi khi xử lý file Excel: {e}")
        return {}
        
def escape_xml_text(text):
    if text:
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        text = text.replace("'", '&apos;')
    return text or ''

def get_schedule_for_channel(channel_id, date_str, uuid_str):
    url = f"https://apigw.mytv.vn/api/v1/channel/{channel_id}/schedule?date={date_str}&uuid={uuid_str}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        schedule = []
        if 'data' in data and 'schedule' in data['data'] and isinstance(data['data']['schedule'], list):
            for item in data['data']['schedule']:
                raw_title = item.get('title', '').replace('\r', '').replace('\t', '').strip()
                processed_title = format_title_string(raw_title) # Áp dụng định dạng
                title = escape_xml_text(processed_title) # Sau đó escape XML

                start_time_str = f"{date_str} {item.get('time')}"
                try:
                    start_datetime = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M")
                    local_timezone = pytz.timezone('Asia/Ho_Chi_Minh')
                    local_datetime = local_timezone.localize(start_datetime)
                    start_formatted = local_datetime.strftime('%Y%m%d%H%M%S +0700')
                    schedule.append({
                        'start': start_formatted,
                        'title': title,
                        'local_start_time': local_datetime
                    })
                except ValueError:
                    print(f"Không thể chuyển đổi thời gian: {start_time_str}")
        print(f"Lịch phát sóng cho kênh {channel_id}:", schedule)
        return schedule
    except requests.exceptions.RequestException as e:
        print(f"Lỗi khi lấy lịch phát sóng cho kênh {channel_id}: {e}")
        return []
    except json.JSONDecodeError:
        print(f"Lỗi khi giải mã JSON lịch phát sóng cho kênh {channel_id}.")
        return []

def create_epg_xml(channels_data, output_path, api_uuid_str):
    now = datetime.now(timezone.utc).astimezone(pytz.timezone('Asia/Ho_Chi_Minh'))
    tv_date_str = now.strftime("%d-%m-%Y")
    generator_info_time = now.strftime("%H:%M:%S - %d/%m/%Y")

    xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<tv date="{tv_date_str}" source-info-name="Ngan Phuc" generator-info-name="Cap nhat luc {generator_info_time}">
"""

    # Tạo các thẻ channel
    for channel_id, info in channels_data.items():
        display_name = escape_xml_text(info.get('display-name', ''))
        display_number = escape_xml_text(str(info.get('display-number', ''))) # Chuyển đổi sang chuỗi
        excel_channel_id = info.get('excel_channel_id')
        xml_content += f"""
  <channel id="{excel_channel_id}">
    <display-name lang="vi">{display_name}</display-name>
    <display-number>{display_number}</display-number>
  </channel>
"""

    # Tạo các thẻ programme
    for channel_id, info in channels_data.items():
        excel_channel_id = info.get('excel_channel_id')
        schedule_data = get_schedule_for_channel(info.get('api_channel_id'), date.today().strftime("%Y-%m-%d"), api_uuid_str)
        if schedule_data:
            for i, program in enumerate(schedule_data):
                start_time = program['start']
                title = escape_xml_text(program['title'])
                end_time = ''
                length_minutes = ''

                if i + 1 < len(schedule_data):
                    next_program_local_start_time = schedule_data[i + 1]['local_start_time']
                    current_program_local_start_time = program['local_start_time']
                    duration = next_program_local_start_time - current_program_local_start_time
                    length_minutes = int(duration.total_seconds() / 60)
                    end_datetime_local = current_program_local_start_time + timedelta(minutes=length_minutes)
                    end_time = end_datetime_local.strftime('%Y%m%d%H%M%S +0700')
                elif schedule_data:
                    length_minutes = 30
                    end_datetime_local = program['local_start_time'] + timedelta(minutes=length_minutes)
                    end_time = end_datetime_local.strftime('%Y%m%d%H%M%S +0700')

                xml_content += f"""
  <programme start="{start_time}" stop="{end_time}" channel="{excel_channel_id}">
    <title lang="vi">{title}</title>
    <length lang="vi">Chương trình này có thời lượng {length_minutes} phút</length>
  </programme>
"""

    xml_content += "</tv>"

    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        print(f"Đã tạo file epg.xml thành công tại: {output_path}")
    except Exception as e:
        print(f"Lỗi khi ghi file epg.xml: {e}")


# Đường dẫn file Excel và file XML đầu ra
excel_file_path = r"G:\Dropbox\IPTV_OPENWRT\channel_list.xlsx"
output_xml_path = r"G:\Dropbox\IPTV_OPENWRT\epg.xml"
api_uuid = str(uuid.uuid4())

# Lấy danh sách kênh từ API
api_channels = get_channel_list_from_api(api_uuid)

# So sánh danh sách kênh và lấy các kênh khớp
matched_channels = compare_channel_lists(excel_file_path, api_channels)

# Tạo file epg.xml dựa trên danh sách kênh khớp
create_epg_xml(matched_channels, output_xml_path, api_uuid)