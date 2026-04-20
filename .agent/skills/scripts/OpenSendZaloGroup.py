import os
import time
import sys
import subprocess
import traceback
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---------------------------------------------------------
# CẤU HÌNH CƠ BẢN
# ---------------------------------------------------------
if getattr(sys, 'frozen', False):
    # Nếu đang chạy từ file .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Nếu đang chạy từ script python
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CREDENTIALS_FILE = os.path.join(BASE_DIR, "gen-lang-client-0450618162-54ea7d476a02.json")
SPREADSHEET_ID = '1SFAr1CFMzMPQXFToZEAwA2U1FaHpeCQqv7CyMa-f-0w'
WORKSHEET_INDEX = 0 # Worksheet đầu tiên
IMAGES_DIR = os.path.join(BASE_DIR, 'images')
CHROME_PROFILE_DIR = os.path.join(BASE_DIR, 'zalo-chrome-profile')

# ---------------------------------------------------------
# UNICODE FIX CHO CMD WINDOWS
# ---------------------------------------------------------
def safe_print(*args, **kwargs):
    text = " ".join(str(a) for a in args)
    try:
        sys.stdout.buffer.write(text.encode('utf-8') + b'\n')
    except Exception:
        print(*args, **kwargs)

# ---------------------------------------------------------
# KẾT NỐI GOOGLE SHEETS
# ---------------------------------------------------------
def get_gspread_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client

# ---------------------------------------------------------
# KHỞI TẠO SELENIUM (CHROME)
# ---------------------------------------------------------
def build_driver():
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={CHROME_PROFILE_DIR}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver

# ---------------------------------------------------------
# CÁC HÀM TƯƠNG TÁC ZALO
# ---------------------------------------------------------
def search_and_click_group(driver, wait, group_name):
    try:
        # Bước 1: Click "Danh bạ" bên trái
        danh_ba_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-translate-title='STR_TAB_CONTACT']")))
        danh_ba_tab.click()
        time.sleep(0.5)
        
        # Bước 2: Click "Danh sách nhóm và cộng đồng"
        group_list_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@data-translate-inner='STR_CONTACT_TAB_GROUP_COMMUNITY_LIST']/..")))
        group_list_tab.click()
        time.sleep(0.5)
        
        # Bước 3: Gõ vào ô "Tìm kiếm..." thuộc tính năng Danh sách nhóm
        search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@data-translate-placeholder='STR_CONTACT_TAB_SEARCH_GROUP']")))
        search_input.clear()
        search_input.send_keys(group_name)
        time.sleep(2) # Đợi kết quả hiển thị
        
        # Bước 4: Click kết quả hiển thị đúng tên nhóm
        xpath_result = f"//div[contains(@class, 'contact-item-v2-wrapper')]//span[@class='name' and text()='{group_name}']"
        try:
            first_result = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_result)))
            # Click vào wrapper cha cao nhất chứ không click trực tiếp span name cho an toàn
            # Hoặc click thẳng span cũng được vì Selenium v4 sẽ auto click vào element
            first_result.click()
        except TimeoutException:
            # Fallback nếu tên có khác biệt nhỏ (do spacing), thì lấy kết quả đầu tiên bấy kỳ
            fallback_xpath = "//div[contains(@class, 'contact-item-v2-wrapper')]"
            first_result = wait.until(EC.element_to_be_clickable((By.XPATH, fallback_xpath)))
            first_result.click()
            
        time.sleep(1) # Đợi Zalo mở cửa sổ chat
        return True
    except TimeoutException:
        safe_print(f"[-] Không tìm thấy nhóm hoặc contact: {group_name}")
        return False
    except Exception as e:
        safe_print(f"[-] Lỗi click kết quả tìm kiếm: {e}")
        return False

def send_message(driver, wait, message, image_keyword):
    # Tìm editor
    editor = wait.until(EC.presence_of_element_located((By.ID, "richInput")))
    
    # Xử lý gửi ảnh (nếu có)
    if image_keyword and str(image_keyword).strip() != "" and str(image_keyword).strip() != "nan":
        img_name = str(image_keyword).strip()
        if not img_name.endswith('.png') and not img_name.endswith('.jpg'):
            img_name += '.png'
        
        abs_img_path = os.path.join(IMAGES_DIR, img_name)
        if os.path.exists(abs_img_path):
            safe_print(f"[*] Đang nạp ảnh: {abs_img_path} vào clipboard...")
            cmd = f'powershell -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile(\'{abs_img_path}\'))"'
            subprocess.run(cmd, shell=True)
            time.sleep(1) # wait clipboard logic
            editor.send_keys(Keys.CONTROL, 'v')
            time.sleep(2) # Chờ ảnh load
            
            # Stale Element Fix (DOM refreshed sau khi dán ảnh)
            editor = wait.until(EC.presence_of_element_located((By.ID, "richInput")))
        else:
            safe_print(f"[-] CẢNH BÁO: Không tìm thấy file ảnh {abs_img_path}")
    
    # Gửi text message
    if message and str(message).strip() != "" and str(message).strip() != "nan":
        safe_print(f"[*] Đang nạp nội dung văn bản vào clipboard...")
        # Sử dụng tạm file để đảm bảo encoding UTF-8 khi nạp vào Clipboard qua PowerShell
        temp_txt_path = os.path.join(BASE_DIR, "temp_msg.txt")
        try:
            with open(temp_txt_path, "w", encoding="utf-8") as f:
                f.write(str(message))
            
            # PowerShell command để đọc file UTF-8 và đưa vào Clipboard
            cmd_text = f'powershell -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetText([System.IO.File]::ReadAllText(\'{temp_txt_path}\', [System.Text.Encoding]::UTF8))"'
            subprocess.run(cmd_text, shell=True)
            time.sleep(0.5)
            
            editor = wait.until(EC.presence_of_element_located((By.ID, "richInput")))
            editor.send_keys(Keys.CONTROL, 'v')
            time.sleep(1)
            
            if os.path.exists(temp_txt_path):
                os.remove(temp_txt_path)
        except Exception as e:
            safe_print(f"[-] Lỗi nạp text vào clipboard: {e}")
            # Fallback nếu clipboard lỗi thì mới dùng send_keys
            editor.send_keys(str(message))
            
    # Nhấn Enter để gửi toàn bộ (cả ảnh và chữ đã dán)
    editor = wait.until(EC.presence_of_element_located((By.ID, "richInput")))
    editor.send_keys(Keys.ENTER)
    time.sleep(1.5)

# ---------------------------------------------------------
# CHƯƠNG TRÌNH CHÍNH
# ---------------------------------------------------------
def main():
    if not os.path.exists(CREDENTIALS_FILE):
        safe_print(f"LỖI TÀI KHOẢN: Không tìm thấy {CREDENTIALS_FILE}")
        return
        
    if not os.path.exists(IMAGES_DIR):
        os.makedirs(IMAGES_DIR)

    safe_print("[*] Đang kết nối Google Sheets...")
    try:
        gc = get_gspread_client()
        sheet = gc.open_by_key(SPREADSHEET_ID).get_worksheet(WORKSHEET_INDEX)
        
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
    except Exception as e:
        safe_print("[-] Lỗi khi kết nối Google Sheets:", e)
        return
        
    if df.empty:
        safe_print("[-] File script rỗng hoặc không tải được dữ liệu.")
        return
        
    # Check if 'Status' column exists
    if 'Status' not in df.columns:
        safe_print("[-] Lỗi: Không tìm thấy cột 'Status' trong sheet.")
        return

    unapproved_rows = df[df['Status'] == 'UNAPPROVED']
    
    if unapproved_rows.empty:
        safe_print("[*] Không có nhóm nào có Status là UNAPPROVED. Dừng chạy.")
        return
    
    safe_print(f"[*] Tìm thấy {len(unapproved_rows)} dòng UNAPPROVED. Đang khởi động Zalo Web...")
    
    driver = build_driver()
    wait = WebDriverWait(driver, 20)
    wait_login = WebDriverWait(driver, 300) # Cho phép quét mã QR lần đầu tới 5 phút
    
    try:
        driver.get("https://chat.zalo.me")
        safe_print("[*] Đang chờ Zalo load xong giao diện (Quét mã QR nếu là lần đầu tiên)...")
        wait_login.until(EC.presence_of_element_located((By.ID, "contact-search-input")))
        time.sleep(3)
        
        for index, row in unapproved_rows.iterrows():
            group_name = str(row.get('Group', ''))
            message = str(row.get('Message', ''))
            image_keyword = str(row.get('Image', ''))
            
            safe_print(f"\n----------------------------------------")
            safe_print(f"[*] Xử lý dòng {index + 2}: Gửi tin nhắn đến Group/SĐT '{group_name}'")
            
            success = search_and_click_group(driver, wait, group_name)
            if success:
                try:
                    send_message(driver, wait, message, image_keyword)
                    safe_print(f"[+] Gửi thành công! Cập nhật trạng thái thành APPROVED.")
                    
                    status_col_index = df.columns.get_loc("Status") + 1
                    sheet.update_cell(index + 2, status_col_index, 'APPROVED')
                except Exception as ex:
                    safe_print(f"[-] Lỗi trong quá trình gửi tin nhắn cho '{group_name}': {ex}")
                    sheet.update_cell(index + 2, df.columns.get_loc("Status") + 1, 'ZALO_FAILED')
            else:
                sheet.update_cell(index + 2, df.columns.get_loc("Status") + 1, 'ZALO_FAILED')
                
            time.sleep(2)
            
    except Exception as e:
        safe_print("[-] Lỗi không mong muốn trong quá trình chạy:", e)
        traceback.print_exc()
        
    finally:
        safe_print("[*] Hoàn thành luồng. Đang đóng trình duyệt sau 5s...")
        time.sleep(5)
        driver.quit()

if __name__ == "__main__":
    main()
