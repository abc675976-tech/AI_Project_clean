import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import ctypes
import ssl
import re
import easyocr
import os
import threading
import sys
from openpyxl.drawing.image import Image as OpenpyxlImage

# 確保路徑正確：無論是開發還是執行 .exe 都能找到檔案
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# 所有的檔案讀寫都必須加上 base_path
invoice_path = os.path.join(base_path, "invoice_data.csv")
learn_path = os.path.join(base_path, "learn.csv")

# 設定圖表中文字型 (解決別台電腦方塊字問題)
plt.rcParams['font.family'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

# 取得目前這個程式檔案所在的資料夾路徑
CSV_PATH = invoice_path

#解決 SSL 憑證
ssl._create_default_https_context = ssl._create_unverified_context

#解決 DPI 模糊問題
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

#初始化 AI 模型
reader = easyocr.Reader(['ch_tra', 'en'])

#存資料
try:
    invoice_data = pd.read_csv(CSV_PATH)
except:
    invoice_data = pd.DataFrame(columns=["號碼","金額","類別", "original_text"])

def get_winning_numbers():
    return {"special_prize":"87510041",
            "grand_prize":"32220522",
            "first_prize":["21677046","44662410","31262513"]}

# ----------------------------
# 讀取照片
# ----------------------------
def extract_invoice_info(results):
    invoice_num = ""
    total_amount = ""

    for (bbox, text, prob) in results:
        # 抓 8 位數字
        num_match = re.search(r'\d{8}', text)
        if num_match and not invoice_num:
            invoice_num = num_match.group()

        # 抓總計金額 (過濾出數字)
        if "總計" in text or "Total" in text:
            amount_match = re.search(r'\d+', text)      
            if amount_match:
                total_amount = amount_match.group()
    return invoice_num, total_amount

def clean_ocr_text(text):
    """
    清洗 OCR 抓到的文字，只保留數字，並處理常見的誤判字
    """
    # 處理常見誤判：把大寫 O 改回 0，把小寫 l 或大寫 I 改回 1
    mapping = {'O': '0', 'o': '0', 'l': '1', 'I': '1', 'S': '5', 's': '5'}
    for char, replacement in mapping.items():
        text = text.replace(char, replacement)

    # 使用正則表達式只留下數字
    return re.sub(r'[^0-9]', '', text)

def auto_find_store_and_category(all_ocr_results):
    global learning_db # 這是你從 learn.csv 讀進來的字典
    print("成功進入確認畫面！") # 如果終端機沒印出這行，代表卡在上一動
    clear_frame()

    found_store = ""

    # 遍歷所有 OCR 辨識出的字串
    for res in all_ocr_results:
        text = res[1] # 取得文字內容
        # 比對字典裡的每一個店名
        for store_name in learning_db.keys():
            if str(store_name) in str(text): # 如果 OCR 結果包含已知店名
                found_store = store_name
                return found_store

    return found_store

# --- 點擊按鈕後的執行動作 ---
def run_ai_ocr():
    # 1. 讓使用者選照片

    file_path = filedialog.askopenfilename(title="選擇發票照片", filetypes=[("圖片", "*.jpg *.png *.jpeg")])
    if not file_path:
        return

    # 2. 切換到等待畫面
    show_loading_screen()

    # 3. 開啟新執行緒跑 AI，避免主畫面卡死

    def ocr_thread():
        results = reader.readtext(file_path)
        detected_num, detected_amount = extract_invoice_info(results)
        # 辨識完後，回到主執行緒更新 UI
        root.after(0, lambda: show_ai_input_step(detected_num, detected_amount, results))

    threading.Thread(target=ocr_thread, daemon=True).start()

def show_loading_screen():
    clear_frame()
    status.config(text="") # 清除之前的錯誤訊息
    
    # 顯示等待文字
    tk.Label(frame, text="✨ AI 正在努力辨識中...", 
             font=(FONT_MAIN, 16, "bold"), fg="#F1C40F", bg=BG_COLOR).pack(pady=40)
    
    tk.Label(frame, text="請稍等幾秒，正在掃描發票資訊...", 
             font=(FONT_MAIN, 11), fg="white", bg=BG_COLOR).pack(pady=10)
    
    # 加入一個簡單的進度條視覺效果（選配）
    loading_label = tk.Label(frame, text="🔍 📄 ⏳", font=(FONT_MAIN, 24), bg=BG_COLOR)
    loading_label.pack(pady=20)

# ----------------------------
# AI 掃描後的確認與分類流程
# ----------------------------

def show_ai_input_step(num, amount, all_results):
    clear_frame()
    status.config(text="") # 清除可能的舊錯誤訊息

    # 1. 執行自動找店名與數據清洗
    # 使用我們之前寫的工具：如果 learn.csv 有這家店，直接抓出來
    auto_store = auto_find_store_and_category(all_results)
    clean_num = clean_ocr_text(num)[:8] # 確保號碼乾淨且只有8位
    clean_amt = clean_ocr_text(amount)

    tk.Label(frame, text="確認發票資訊", font=(FONT_MAIN, 18, "bold"), fg="white", bg=BG_COLOR).pack(pady=20)

    # --- 號碼輸入欄位 ---
    tk.Label(frame, text="發票號碼 (8位數字):", fg="white", bg=BG_COLOR).pack()
    num_entry = tk.Entry(frame, font=(FONT_MAIN, 14), justify='center')
    num_entry.insert(0, clean_num)
    num_entry.pack(pady=5)

    # --- 金額輸入欄位 ---
    tk.Label(frame, text="發票金額:", fg="white", bg=BG_COLOR).pack()
    amount_entry = tk.Entry(frame, font=(FONT_MAIN, 14), justify='center')
    amount_entry.insert(0, clean_amt)
    amount_entry.pack(pady=5)

    # --- 店名輸入欄位 (關鍵：這裡自動帶入 AI 辨識到的店名) ---
    tk.Label(frame, text="辨識到的店名 (若正確請直接按確認):", fg="#F1C40F", bg=BG_COLOR).pack()
    store_entry = tk.Entry(frame, font=(FONT_MAIN, 14), justify='center', fg="#F1C40F")
    store_entry.insert(0, auto_store)
    store_entry.pack(pady=5)

    # 2. 定義內部的跳轉邏輯 (必須放在按鈕產生之前)
    def go_to_category():
        global current
        final_n = num_entry.get().strip()
        final_a = amount_entry.get().strip()
        final_s = store_entry.get().strip()

        # 基礎驗證
        if not (final_n.isdigit() and len(final_n) == 8):
            status.config(text="❌ 號碼須為 8 位數字", fg="#FF7675")
            return
        
        try:
            val_a = float(final_a)
        except ValueError:
            status.config(text="❌ 金額請輸入數字", fg="#FF7675")
            return

        # 存入全域字典 current，讓後面的 step_ai() 抓得到
        current["號碼"] = final_n
        current["金額"] = val_a
        text = final_s
        
        if not text:
            status.config(text="❌ 請輸入店名以進行分類", fg="#FF7675")
            return
            
        current["original_text"] = text
        result = ai_classify(text)
        
        # 根據 AI 分類結果跳轉
        if result is None:
            step_ai_manual_choice(text) # 情況 1: 進入手動選擇
        else:
            step_ai_result_2(result, text)

    # 3. 產生按鈕 (現在 Python 找得到 go_to_category 了)
    create_button(frame, "確認無誤，進行 AI 分類", go_to_category).pack(pady=20)

def step_ai_manual_choice(text):
    clear_frame()
    tk.Label(frame, text="請選擇消費類別", font=(FONT_MAIN, 14, "bold"), fg="white", bg=BG_COLOR).pack(pady=15)

    grid = tk.Frame(frame, bg=BG_COLOR)
    grid.pack()
    for i, c in enumerate(["食","衣","住","行","育","樂"]):
        btn = create_button(grid, c, lambda cat=c: [current.update({"類別":cat}), learning_db.update({text:cat}), save_learning(), step4_1()], width=10, height=2)
        btn.grid(row=i//2, column=i%2, padx=10, pady=10)

    create_button(frame, "返回", lambda: show_ai_input_step(current.get("號碼", ""), str(current.get("金額", "")), [])).pack(pady=10)

#給類別選完返回自動分類結果
def step_ai_manual_choice_2(result, text):
    clear_frame()
    tk.Label(frame, text="請選擇消費類別", font=(FONT_MAIN, 14, "bold"), fg="white", bg=BG_COLOR).pack(pady=15)

    grid = tk.Frame(frame, bg=BG_COLOR)
    grid.pack()
    for i, c in enumerate(["食","衣","住","行","育","樂"]):
        btn = create_button(grid, c, lambda cat=c: [current.update({"類別":cat}), learning_db.update({text:cat}), save_learning(), step4_1()], width=10, height=2)
        btn.grid(row=i//2, column=i%2, padx=10, pady=10)

    create_button(frame, "返回", lambda: step_ai_result_2(result, text)).pack(pady=10)

# OCR之分類結果的返回畫面
def step_ai_result_2(result, text):
    clear_frame()
    tk.Label(frame, text="AI 自動分類結果", font=(FONT_MAIN, 16, "bold"), fg="white", bg=BG_COLOR).pack(pady=10)
    tk.Label(frame, text=f"店名: {text}\n預測類別: {result}", font=(FONT_MAIN, 12), fg="#3498DB", bg=BG_COLOR).pack(pady=15)

    btn_frame = tk.Frame(frame, bg=BG_COLOR)
    btn_frame.pack(pady=10)

    # 情況 1: 手動選擇
    create_button(btn_frame, "手動選擇", lambda: step_ai_manual_choice_2(result, text), width=12).grid(row=0, column=0, padx=5)

    # 情況 2: 確認正確 -> 直接進中獎畫面 (同步兩個 CSV)
    def confirm_and_save():
        current["類別"] = result
        learning_db[text] = result # 更新學習字典
        save_learning()            # 同步到 learn.csv
        step4_1()                    # 同步到 invoice_data.csv 並顯示中獎

    create_button(btn_frame, "確認正確", confirm_and_save, width=12).grid(row=0, column=1, padx=5)
    create_button(frame, "返回修改資訊", lambda: show_ai_input_step(current.get("號碼", ""), str(current.get("金額", "")), [])).pack(pady=10)

def final_save(num, amount, cat):
    global invoice_data
    
    # 強制轉換型態，避免比對出錯
    num_str = str(num)
    
    # 建立新資料 (包含原本的 original_text 欄位)
    new_row = pd.DataFrame([[num_str, float(amount), cat, current.get("original_text","")]], 
                           columns=["號碼", "金額", "類別", "original_text"])
    
    # 確保原本的資料號碼也是字串
    invoice_data["號碼"] = invoice_data["號碼"].astype(str)
    
    # 串接並存檔 (指名存到原本的檔案)
    invoice_data = pd.concat([invoice_data, new_row], ignore_index=True)
    invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")
    
    # 回主畫面並給綠色提示
    show_main()
    status.config(text=f"✅ 成功存入！號碼: {num_str}", fg="#55E6C1")

# ----------------------------
# AI學習資料（讀取）
# ----------------------------
try:
    learn_df = pd.read_csv(learn_path)
    learning_db = dict(zip(learn_df["text"], learn_df["category"]))
except:
    learning_db={}

current = {"號碼":"", "金額":0, "類別":""}

def save_learning():
    df = pd.DataFrame(list(learning_db.items()), columns=["text","category"])
    df.to_csv("learn.csv", index=False)

# ----------------------------
# AI簡單分類（關鍵字版）
# ----------------------------
def ai_classify(text):
    text = text.lower()

    if text in learning_db:
        return learning_db[text]

    food = ["食","吃","餐","麥當勞","飲料","早餐","晚餐","便當","薯條","雞塊","壽司","水果"]
    cloth = ["衣","褲","鞋","uniqlo","zara","襪","包"]
    live = ["住","房","租","旅館","hotel","電視","飯店","民宿"]
    move = ["行","捷運","公車","uber","油","高鐵","腳踏車","計程車","taxi","火車"]
    edu = ["育","書","課","補習","學","課程"]
    fun = ["樂","電影","遊樂","netflix","遊戲","ktv","桌遊","酒","動漫","化妝品"]

    for k in food:
        if k in text: return "食"
    for k in cloth:
        if k in text: return "衣"
    for k in live:
        if k in text: return "住"
    for k in move:
        if k in text: return "行"
    for k in edu:
        if k in text: return "育"
    for k in fun:
        if k in text: return "樂"

    return None

# ----------------------------
# 主視窗（置中佈局設定）
# ----------------------------
root = tk.Tk()
root.title("發票分析系統")

w, h = 600, 650
ws, hs = root.winfo_screenwidth(), root.winfo_screenheight()
x, y = (ws-w)//2, (hs-h)//2
root.geometry(f"{w}x{h}+{x}+{y}")

# 設定深色背景，看起來最舒服
BG_COLOR = "#2C3E50" 
root.configure(bg=BG_COLOR)

# ----------------------------
# UI樣式（文字改為微軟正黑體）
# ----------------------------
BTN_COLOR = "#3498DB"
BTN_HOVER = "#2980B9"
BTN_TEXT = "white"
FONT_MAIN = "微軟正黑體"

def create_button(parent, text, command, width=20, height=2):
    btn = tk.Button(parent,
                    text=text,
                    command=command,
                    bg=BTN_COLOR,
                    fg=BTN_TEXT,
                    activebackground=BTN_HOVER,
                    activeforeground="white",
                    font=(FONT_MAIN, 10, "bold"),
                    width=width,
                    height=height,
                    bd=0,
                    cursor="hand2")

    btn.bind("<Enter>", lambda e: btn.config(bg=BTN_HOVER))
    btn.bind("<Leave>", lambda e: btn.config(bg=BTN_COLOR))

    return btn

# ----------------------------
# 狀態標籤（改為正黑體且顏色更亮）
status = tk.Label(root, text="", fg="#FF7675", bg=BG_COLOR, font=(FONT_MAIN, 10))
status.pack(pady=10)

# 使用一個容器讓所有內容垂直水平置中
center_container = tk.Frame(root, bg=BG_COLOR)
center_container.place(relx=0.5, rely=0.5, anchor="center")

frame = tk.Frame(center_container, bg=BG_COLOR)
frame.pack()

# ----------------------------
def clear_frame():
    for widget in frame.winfo_children():
        widget.destroy()

# ----------------------------
# 主選單
# ----------------------------

def show_main():
    clear_frame()
    status.config(text="")

    tk.Label(frame, text="發票分析系統",
             font=(FONT_MAIN, 20, "bold"),
             fg="white",
             bg=BG_COLOR).pack(pady=20)
    
    ai_btn = create_button(frame, "📷 AI 拍照對獎", run_ai_ocr)
    ORANGE_COLOR = "#E67E22"
    ORANGE_HOVER = "#D35400"
    ai_btn.config(bg=ORANGE_COLOR)
    # 重新綁定滑鼠事件，確保離開時是回到橘色而不是藍色
    ai_btn.bind("<Enter>", lambda e: ai_btn.config(bg=ORANGE_HOVER))
    ai_btn.bind("<Leave>", lambda e: ai_btn.config(bg=ORANGE_COLOR))
    ai_btn.pack(pady=8)

    create_button(frame, "發票對獎", step1).pack(pady=8)
    create_button(frame, "消費分析", analyze_step1).pack(pady=8)
    create_button(frame, "刪除指定一筆資料", delete_specific).pack(pady=8)
    create_button(frame, "修改指定一筆資料", edit_specific).pack(pady=8)
    create_button(frame, "清除所有資料", confirm_delete_all).pack(pady=8)

# ----------------------------
# Step1
# ----------------------------
def step1():
    clear_frame()
    status.config(text="")

    tk.Label(frame, text="輸入發票號碼(8位)",
             font=(FONT_MAIN, 12), fg="white", bg=BG_COLOR).pack(pady=10)

    tk.Label(frame, text="（請切換英文/數字輸入）",
             font=(FONT_MAIN, 9), fg="#BDC3C7", bg=BG_COLOR).pack()

    entry = tk.Entry(frame, width=25, font=(FONT_MAIN, 12), justify='center')
    entry.pack(pady=15)
    entry.focus_set()

    def next():
        num = entry.get()
        if len(num)==8 and num.isdigit():
            current["號碼"] = num
            step2()
        else:
            status.config(text="❌ 發票號碼錯誤")

    create_button(frame, "下一步", next).pack(pady=5)
    create_button(frame, "返回主選單", show_main).pack(pady=5)

# ----------------------------
# Step2
# ----------------------------
def step2():
    clear_frame()
    status.config(text="")

    tk.Label(frame, text=f"號碼: {current['號碼']}",
             font=(FONT_MAIN, 11), fg="#BDC3C7", bg=BG_COLOR).pack()

    tk.Label(frame, text="輸入金額",
             font=(FONT_MAIN, 12), fg="white", bg=BG_COLOR).pack(pady=10)

    entry = tk.Entry(frame, width=25, font=(FONT_MAIN, 12), justify='center')
    entry.pack(pady=15)
    entry.focus_set()

    def next():
        try:
            raw_value=entry.get().strip()
            if not raw_value:
                status.config(text="❌ 金額格式錯誤", fg="#FF7675")
                return
            amt = float(raw_value)
            if amt < 0:
                status.config(text="❌ 金額須為正數", fg="#FF7675")
                return
            current["金額"] = amt
            step_ai()
        except:
            status.config(text="❌ 請輸入數字", fg="#FF7675")

    create_button(frame, "下一步", next).pack(pady=5)
    create_button(frame, "返回", step1).pack(pady=5)

def step_ai():
    clear_frame()
    status.config(text="")

    tk.Label(frame, text="輸入店名或商品名稱",
             font=(FONT_MAIN, 12, "bold"),
             fg="white", bg=BG_COLOR).pack(pady=5)
    tk.Label(frame, text="（AI自動分類）",
             font=(FONT_MAIN, 10), fg="#BDC3C7", bg=BG_COLOR).pack(pady=5)

    entry = tk.Entry(frame, width=25, font=(FONT_MAIN, 12), justify='center')
    entry.pack(pady=15)
    entry.focus_set()

    def predict():
        text = entry.get()
        if text == "":
            status.config(text="❌ 請輸入內容")
            return
        
        current["original_text"] = text
        result = ai_classify(text)

        if result is None:
            step_ai_unknown(text)
        else:
            current["類別"] = result
            step_ai_result(result, text)

    create_button(frame, "AI分類", predict).pack(pady=5)
    create_button(frame, "返回", step2).pack(pady=5)

def step_ai_unknown(text):
    clear_frame()
    status.config(text="")

    tk.Label(frame, text="AI無法判斷,請手動選擇",
             font=(FONT_MAIN, 14, "bold"),
             fg="#FAB1A0",
             bg=BG_COLOR).pack(pady=15)

    categories = ["食","衣","住","行","育","樂"]

    def choose(c):
        global learning_db
        current["類別"] = c
        learning_db[text] = c
        save_learning()
        step4()

    # 類別按鈕改用網格置中顯示
    grid_btn = tk.Frame(frame, bg=BG_COLOR)
    grid_btn.pack()
    for i, c in enumerate(categories):
        btn = create_button(grid_btn, c, lambda c=c: choose(c), width=8, height=1)
        btn.grid(row=i//2, column=i%2, padx=5, pady=5)

    create_button(frame, "返回", step_ai).pack(pady=15)

def step_ai_result(result, text):
    clear_frame()
    tk.Label(frame, text="AI 自動分類結果", font=(FONT_MAIN, 16, "bold"), fg="white", bg=BG_COLOR).pack(pady=10)
    tk.Label(frame, text=f"店名: {text}\n預測類別: {result}", font=(FONT_MAIN, 12), fg="#3498DB", bg=BG_COLOR).pack(pady=15)

    btn_frame = tk.Frame(frame, bg=BG_COLOR)
    btn_frame.pack(pady=10)

    # 情況 1: 手動選擇
    create_button(btn_frame, "手動選擇", lambda: step_ai_unknown(text), width=12).grid(row=0, column=0, padx=5)

    # 情況 2: 確認正確 -> 直接進中獎畫面 (同步兩個 CSV)
    def confirm_and_save():
        current["類別"] = result
        learning_db[text] = result # 更新學習字典
        save_learning()            # 同步到 learn.csv
        step4()                    # 同步到 invoice_data.csv 並顯示中獎

    create_button(btn_frame, "確認正確", confirm_and_save, width=12).grid(row=0, column=1, padx=5)
    create_button(frame, "返回修改資訊", step_ai).pack(pady=10)

# ----------------------------
# Step3 類別（大按鈕）
# ----------------------------
def step3(text_from_ai=None):
    clear_frame()
    status.config(text="")

    tk.Label(frame, text="選擇消費類別",
             font=(FONT_MAIN, 14, "bold"),
             fg="white",
             bg=BG_COLOR).pack(pady=15)

    grid = tk.Frame(frame, bg=BG_COLOR)
    grid.pack()

    categories = ["食","衣","住","行","育","樂"]

    def choose(c):
        current["類別"] = c
        if text_from_ai:
            learning_db[text_from_ai] = c
            save_learning()         
        step4()

    for i, c in enumerate(categories):
        btn = create_button(grid, c, lambda c=c: choose(c), width=10, height=2)
        btn.grid(row=i//2, column=i%2, padx=10, pady=10)

    create_button(frame, "返回", step2).pack(pady=15)

# ----------------------------
# Step4 結果
# ----------------------------
def step4():
    clear_frame()
    status.config(text="")

    num = current["號碼"]
    winning = get_winning_numbers()
    ns=winning["special_prize"]
    n1=winning["grand_prize"]
    n2=winning["first_prize"]
    prizes={7:40000,6:10000,5:4000,4:1000,3:200}

    result="未中獎"

    if num==ns:
        result="中1000萬!"
    elif num==n1:
        result="中200萬!"
    else:
        for i in n2:
            if num==i:
                result="中20萬!"
                break
        else:
            for i in n2:
                for n in range(7,2,-1):
                    if num[-n:]==i[-n:]:
                        result=f"中{n}碼,獲得{prizes[n]}元"
                        break

    global invoice_data
    new = pd.DataFrame([[current["號碼"], current["金額"], current["類別"], current.get("original_text","")]], 
                       columns=["號碼", "金額", "類別", "original_text"])
    
    invoice_data["號碼"] = invoice_data["號碼"].astype(str)
    new["號碼"] = new["號碼"].astype(str)

    invoice_data = pd.concat([invoice_data, new], ignore_index=True)
    invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")

    tk.Label(frame, text=f"結果: {result}",
             font=(FONT_MAIN, 16, "bold"),
             fg="#F1C40F",
             bg=BG_COLOR).pack(pady=20)

    tk.Label(frame, text=f"已儲存，共 {len(invoice_data)} 筆資料",
             font=(FONT_MAIN, 10), fg="#BDC3C7", bg=BG_COLOR).pack()

    create_button(frame, "繼續輸入", step1).pack(pady=8)
    create_button(frame, "回主選單", show_main).pack(pady=8)

#沒有繼續輸入選項
def step4_1():
    clear_frame()
    status.config(text="")

    num = current["號碼"]
    winning = get_winning_numbers()
    ns=winning["special_prize"]
    n1=winning["grand_prize"]
    n2=winning["first_prize"]
    prizes={7:40000,6:10000,5:4000,4:1000,3:200}

    result="未中獎"

    if num==ns:
        result="中1000萬!"
    elif num==n1:
        result="中200萬!"
    else:
        for i in n2:
            if num==i:
                result="中20萬!"
                break
        else:
            for i in n2:
                for n in range(7,2,-1):
                    if num[-n:]==i[-n:]:
                        result=f"中{n}碼 {prizes[n]}元"
                        break

    global invoice_data
    new = pd.DataFrame([[current["號碼"], current["金額"], current["類別"], current.get("original_text","")]], 
                       columns=["號碼", "金額", "類別", "original_text"])
    
    invoice_data["號碼"] = invoice_data["號碼"].astype(str)
    new["號碼"] = new["號碼"].astype(str)

    invoice_data = pd.concat([invoice_data, new], ignore_index=True)
    invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")

    tk.Label(frame, text=f"結果: {result}",
             font=(FONT_MAIN, 16, "bold"),
             fg="#F1C40F",
             bg=BG_COLOR).pack(pady=20)

    tk.Label(frame, text=f"已儲存，共 {len(invoice_data)} 筆資料",
             font=(FONT_MAIN, 10), fg="#BDC3C7", bg=BG_COLOR).pack()
    
    create_button(frame, "回主選單", show_main).pack(pady=8)
# ----------------------------
# AI分析
# ----------------------------
def export_to_excel():
    global invoice_data, chart_status_label
    if invoice_data.empty:
        if 'chart_status_label' in globals():
            chart_status_label.config(text="❌ 目前沒有資料可以匯出", fg="#FF7675")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="儲存消費分析報表"
    )

    if file_path:
        try:
            # 1. 準備數據 (保留 original_text 也就是店名)
            export_df = invoice_data.copy()
            # 關鍵：確保號碼是字串，金額是數字
            export_df["號碼"] = export_df["號碼"].astype(str)
            export_df["金額"] = pd.to_numeric(export_df["金額"], errors='coerce')
            
            # 2. 製作統計摘要 (加上 as_index=False 確保「類別」不會變成空白索引)
            summary = export_df.groupby("類別", as_index=False)["金額"].sum()
            summary.columns = ["消費類別", "總金額 (TWD)"]

            # 2. 產出圖表暫存檔 (圓餅、長條、折線)
            fig_temp, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(6, 12))
            # 圓餅
            ax1.pie(summary["總金額 (TWD)"], labels=summary["消費類別"], autopct='%1.1f%%', startangle=140)
            ax1.set_title("消費比例分析")
            # 長條
            ax2.bar(summary["消費類別"], summary["總金額 (TWD)"], color='#3498DB')
            ax2.set_title("各類別總額統計")
            # 折線
            ax3.plot(range(1, len(export_df)+1), export_df["金額"], marker='o', color='#E67E22')
            ax3.set_title("個人消費趨勢追蹤")
            
            plt.tight_layout()
            chart_path = os.path.join(base_path, "full_report_chart.png")
            fig_temp.savefig(chart_path, dpi=100)
            plt.close(fig_temp)

            # 3. 寫入 Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 分頁 1：消費明細
                export_df.to_excel(writer, sheet_name='消費明細', index=False)
                
                # 分頁 2：視覺化分析
                summary.to_excel(writer, sheet_name='視覺化分析', index=False, startrow=0, startcol=0)
                
                worksheet = writer.sheets['視覺化分析']
                img = OpenpyxlImage(chart_path)
                worksheet.add_image(img, 'A8')

            if os.path.exists(chart_path):
                os.remove(chart_path)
            
            # 更新顯示在圖表下方的提示
            chart_status_label.config(text=f"✅ 報表已匯出：{os.path.basename(file_path)}", fg="#55E6C1")
            
        except Exception as e:
            chart_status_label.config(text=f"❌ 匯出失敗: {str(e)}", fg="#FF7675")

def analyze_step1():
    clear_frame()

    if len(invoice_data) < 5:
        tk.Label(frame, text="❌ 資料筆數不足", 
                 font=(FONT_MAIN, 16, "bold"), fg="#E74C3C", bg=BG_COLOR).pack(pady=10)
        tk.Label(frame, text=f"目前僅有 {len(invoice_data)} 筆\n至少需要 5 筆才能進行分析。", 
                 font=(FONT_MAIN, 11), fg="white", bg=BG_COLOR, justify='center').pack(pady=10)
        
        create_button(frame, "返回主選單", show_main).pack(pady=20)
        return

    cs = invoice_data.groupby("類別")["金額"].sum()
    total = cs.sum()
    top_category = cs.idxmax()
    top_ratio = cs.max() / total

    persona_map = {
        "食": "美食型消費者", "衣": "時尚型消費者", "住": "居家型消費者",
        "行": "奔波型消費者", "育": "進修型消費者", "樂": "及時行樂型消費者"
    }
    p = persona_map.get(top_category, f"{top_category}愛好者") if top_ratio > 0.35 else "均衡型消費者"

    avg = invoice_data["金額"].mean()
    predict = int(avg*10)

    advice = (
        f"主要支出 : {top_category} (佔 {top_ratio:.1%})\n\n"
        f"核心人格 : {p}\n\n"
        f"未來預估 : 下 10 筆消費約 {predict:,} TWD"
    )

    tk.Label(frame, text="📊 消費行為診斷", 
             font=(FONT_MAIN, 16, "bold"), fg="#00d2ff", bg=BG_COLOR).pack(pady=15)

    # 加入深色對話框感
    tk.Label(frame, text=advice, font=(FONT_MAIN, 11),
             bg="#34495E", fg="#C9E622", padx=20, pady=15, 
             justify='center', relief="groove").pack(pady=10)

    create_button(frame, "繼續分析（查看圖表）", show_charts).pack(pady=8)
    create_button(frame, "返回主選單", show_main).pack(pady=8)

# ----------------------------
# 圖表
# ----------------------------
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def show_charts():
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
    plt.rcParams['axes.unicode_minus'] = False

    if invoice_data.empty:
        status.config(text="❌ 目前尚無消費資料", fg="#EA2027")
        return
    
    plot_df = invoice_data.copy()
    plot_df["金額"] = pd.to_numeric(plot_df["金額"], errors='coerce')
    plot_df = plot_df.dropna(subset=["金額"])
    data_sum = plot_df.groupby("類別")["金額"].sum().sort_values(ascending=False)
    plot_df["序號"] = range(1, len(plot_df) + 1)

    clear_frame()
    
    container = tk.Canvas(frame, bg=BG_COLOR, highlightthickness=0, width=600, height=650)
    scrollable_frame = tk.Frame(container, bg=BG_COLOR)
    scrollable_frame.bind("<Configure>", lambda e: container.configure(scrollregion=container.bbox("all")))
    
    container.create_window((300, 0), window=scrollable_frame, anchor="n")
    container.pack(side="left", fill="both", expand=True)

    def _on_mousewheel(event):
        container.yview_scroll(int(-1*(event.delta/120)), "units")
    root.bind_all("<MouseWheel>", _on_mousewheel)

    # 1. 建立畫布
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(5.0, 11), facecolor=BG_COLOR, dpi=100)
    t_font = {'color': 'white', 'fontsize': 11, 'fontweight': 'bold'}

    # (1) 圓餅圖
    ax1.set_position([0.25, 0.72, 0.6, 0.22]) 
    ax1.pie(data_sum.values, labels=data_sum.index, autopct='%1.1f%%', 
            startangle=140, pctdistance=0.7,
            textprops={'color':"white", 'fontsize': 9}, colors=plt.cm.Pastel1.colors)
    for text in ax1.texts:
        if "%" in text.get_text(): text.set_color('black')
    ax1.set_title("消費比例分析", **t_font, pad=10)

    # (2) 長條圖
    ax2.set_position([0.32, 0.40, 0.6, 0.22]) 
    ax2.bar(data_sum.index, data_sum.values, color='#3498DB', width=0.5)
    ax2.set_title("各類別總金額統計", **t_font, pad=15)
    ax2.tick_params(colors='white', labelsize=9)
    if not data_sum.empty: 
        ax2.set_ylim(0, data_sum.max() * 1.3)
    ax2.set_ylabel("金\n額\n\n$\mathit{T}$\n$\mathit{W}$\n$\mathit{D}$", color='white', fontsize=9, rotation=0, labelpad=15, va='center')
    for bar in ax2.patches:
        ax2.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 1, 
                 f'{int(bar.get_height())}', ha='center', va='bottom', color='black', fontsize=8, fontweight='bold')

    # (3) 折線圖
    ax3.set_position([0.32, 0.08, 0.6, 0.22]) 
    ax3.plot(plot_df["序號"], plot_df["金額"], marker='o', color='#E67E22', markersize=5, linewidth=2)
    ax3.fill_between(plot_df["序號"], plot_df["金額"].astype(float), color='#E67E22', alpha=0.15)
    ax3.set_title("個人消費趨勢追蹤", **t_font, pad=15)
    ax3.tick_params(colors='white', labelsize=9)
    if not plot_df["金額"].empty: 
        ax3.set_ylim(0, plot_df["金額"].max() * 1.3)
    ax3.set_ylabel("金\n額\n\n$\mathit{T}$\n$\mathit{W}$\n$\mathit{D}$", color='white', fontsize=9, rotation=0, labelpad=15, va='center')
    ax3.grid(True, linestyle=':', alpha=0.2)

    canvas_plot = FigureCanvasTkAgg(fig, master=scrollable_frame)
    canvas_plot.draw()
    canvas_plot.get_tk_widget().pack(pady=10)

    global chart_status_label
    chart_status_label = tk.Label(scrollable_frame, text="", font=(FONT_MAIN, 10), bg=BG_COLOR)
    chart_status_label.pack(pady=(20, 0), padx=(100, 0))

    # --- 在圖表頁面新增 Excel 匯出按鈕 ---
    excel_btn = create_button(scrollable_frame, "📥 匯出 Excel 完整報表", export_to_excel)
    excel_btn.config(bg="#27ae60") 
    excel_btn.pack(pady=20, padx=(100, 0))

    def go_back():
        root.unbind_all("<MouseWheel>")
        show_main()
    create_button(scrollable_frame, "返回主選單", go_back).pack(pady=20, padx=(100, 0))

# ----------------------------
# 刪除指定資料
# ----------------------------
def delete_specific():
    global invoice_data
    clear_frame() 
    status.config(text="")

    if invoice_data.empty:
        tk.Label(frame, text="⚠️ 目前沒有資料可刪除", 
                 font=(FONT_MAIN, 12), bg=BG_COLOR, fg="#E67E22").pack(pady=20)
        create_button(frame, "返回主選單", show_main).pack()
        return

    tk.Label(frame, text="請選擇要刪除的資料編號",
             font=(FONT_MAIN, 12, "bold"), fg="white", bg=BG_COLOR).pack(pady=10)

    list_frame = tk.Frame(frame, bg=BG_COLOR)
    list_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(list_frame, bg=BG_COLOR, height=250, highlightthickness=0)
    scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview, width=16)
    scroll_inner = tk.Frame(canvas, bg=BG_COLOR)

    scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y",padx=10, pady=5)
    canvas.pack(side="left", fill="both", expand=True)
    

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    for i in range(len(invoice_data)):
        row = invoice_data.iloc[i]
        orig = row.get("original_text", "")
        cat_display = f"{row['類別']} ({orig})" if pd.notna(orig) and orig != "" else row['類別']
        display_text = f" {i+1}. {row['號碼']} | ${int(row['金額'])} | {cat_display}"
        tk.Label(scroll_inner, text=display_text, bg=BG_COLOR, fg="white", font=("Consolas", 10),padx=20).pack(anchor="w")

    bottom_frame = tk.Frame(frame, bg=BG_COLOR)
    bottom_frame.pack(side="bottom", fill="x", pady=10)

    input_box = tk.Frame(bottom_frame, bg=BG_COLOR)
    input_box.pack(pady=5)
    tk.Label(input_box, text="請輸入編號:", font=(FONT_MAIN, 10), fg="white", bg=BG_COLOR).grid(row=0, column=0)
    entry = tk.Entry(input_box, width=8, font=(FONT_MAIN, 10), justify='center')
    entry.grid(row=0, column=1, padx=5)
    entry.focus_set()

    btn_box = tk.Frame(bottom_frame, bg=BG_COLOR)
    btn_box.pack(pady=10)

    def do_delete():
        global invoice_data
        try:
            num = int(entry.get())
            idx = num - 1
            if 0 <= idx < len(invoice_data):
                invoice_data = invoice_data.drop(invoice_data.index[idx]).reset_index(drop=True)
                invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")
                delete_specific()
                status.config(text=f"✅ 已成功刪除第 {num} 筆", fg="#55E6C1")
            else:
                status.config(text=f"❌ 找不到編號 {num}")
        except:
            status.config(text="❌ 請輸入數字編號")

    create_button(btn_box, "確認刪除", do_delete, width=12).grid(row=0, column=0, padx=10)
    create_button(btn_box, "返回主選單", show_main, width=12).grid(row=0, column=1, padx=10)

# ----------------------------
# 清除所有資料
# ----------------------------
def confirm_delete_all():
    global invoice_data
    if invoice_data.empty:
        status.config(text="⚠️ 目前本來就沒有資料喔！", fg="#E67E22")
        return

    clear_frame()
    status.config(text="")

    tk.Label(frame, text="☢️ 危險操作確認 ☢️", 
             font=(FONT_MAIN, 18, "bold"), fg="#E74C3C", bg=BG_COLOR).pack(pady=20)
    
    tk.Label(frame, text="您確定要清空「所有」發票紀錄嗎？\n此動作將無法復原！", 
             font=(FONT_MAIN, 11), fg="white", bg=BG_COLOR, justify='center').pack(pady=10)

    btn_box = tk.Frame(frame, bg=BG_COLOR)
    btn_box.pack(pady=20)

    btn_danger = tk.Button(btn_box, text="確定全部刪除", 
                           command=execute_delete_all,
                           font=(FONT_MAIN, 10, "bold"),
                           bg="#E74C3C", fg="white", 
                           width=15, height=2, bd=0, cursor="hand2")
    btn_danger.grid(row=0, column=0, padx=10)

    create_button(btn_box, "點錯了，返回", show_main, width=15).grid(row=0, column=1, padx=10)

def execute_delete_all():
    global invoice_data
    invoice_data = pd.DataFrame(columns=["號碼", "金額", "類別", "original_text"])
    invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")
    show_main()
    status.config(text="✅ 所有資料已成功清空！", fg="#55E6C1")

# ----------------------------
# 修改資料
# ----------------------------
def edit_specific():
    global invoice_data
    clear_frame()
    status.config(text="")

    if invoice_data.empty:
        tk.Label(frame, text="⚠️ 目前沒有資料可修改", 
                 font=(FONT_MAIN, 12), bg=BG_COLOR, fg="#E67E22").pack(pady=20)
        create_button(frame, "返回主選單", show_main).pack()
        return

    tk.Label(frame, text="請輸入要修改的資料編號", 
             font=(FONT_MAIN, 12, "bold"), fg="white", bg=BG_COLOR).pack(pady=10)

    list_frame = tk.Frame(frame, bg=BG_COLOR)
    list_frame.pack(fill="both", expand=True)
    canvas = tk.Canvas(list_frame, bg=BG_COLOR, height=250, highlightthickness=0)
    scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview, width=16)
    scroll_inner = tk.Frame(canvas, bg=BG_COLOR)
    scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y",padx=10, pady=5)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    for i in range(len(invoice_data)):
        row = invoice_data.iloc[i]
        orig = row.get("original_text", "")
        cat_display = f"{row['類別']} ({orig})" if pd.notna(orig) and orig != "" else row['類別']
        display_text = f" {i+1}. {row['號碼']} | ${int(row['金額'])} | {cat_display}"
        tk.Label(scroll_inner, text=display_text, bg=BG_COLOR, fg="white", font=("Consolas", 10),padx=20).pack(anchor="w")

    bottom_frame = tk.Frame(frame, bg=BG_COLOR)
    bottom_frame.pack(side="bottom", fill="x", pady=10)
    
    input_box = tk.Frame(bottom_frame, bg=BG_COLOR)
    input_box.pack(pady=5)
    tk.Label(input_box, text="修改編號:", font=(FONT_MAIN, 10), fg="white", bg=BG_COLOR).grid(row=0, column=0)
    entry = tk.Entry(input_box, width=8, font=(FONT_MAIN, 10), justify='center')
    entry.grid(row=0, column=1, padx=5)
    entry.focus_set()

    def go_to_edit():
        try:
            num = int(entry.get())
            if 1 <= num <= len(invoice_data):
                status.config(text="")
                edit_detail_screen(num - 1)
            else:
                status.config(text=f"❌ 找不到編號 {num}")
        except:
            status.config(text="❌ 請輸入數字編號")

    btn_box = tk.Frame(bottom_frame, bg=BG_COLOR)
    btn_box.pack(pady=10)
    create_button(btn_box, "開始修改", go_to_edit, width=12).grid(row=0, column=0, padx=10)
    create_button(btn_box, "返回主選單", show_main, width=12).grid(row=0, column=1, padx=10)

def edit_detail_screen(idx):
    clear_frame()
    global invoice_data
    old_row = invoice_data.iloc[idx]
    
    tk.Label(frame, text=f"正在修改第 {idx+1} 筆資料", 
             font=(FONT_MAIN, 14, "bold"), fg="white", bg=BG_COLOR).pack(pady=15)

    tk.Label(frame, text="發票號碼:", font=(FONT_MAIN, 10), fg="white", bg=BG_COLOR).pack()
    en_num = tk.Entry(frame, width=25, font=(FONT_MAIN, 11), justify='center')
    en_num.insert(0, str(old_row["號碼"])) 
    en_num.pack(pady=8)

    tk.Label(frame, text="金額:", font=(FONT_MAIN, 10), fg="white", bg=BG_COLOR).pack()
    en_amt = tk.Entry(frame, width=25, font=(FONT_MAIN, 11), justify='center')
    en_amt.insert(0, str(old_row["金額"])) 
    en_amt.pack(pady=8)

    tk.Label(frame, text="店名/品項:", font=(FONT_MAIN, 10), fg="white", bg=BG_COLOR).pack()
    en_orig = tk.Entry(frame, width=25, font=(FONT_MAIN, 11), justify='center')
    en_orig.insert(0, str(old_row.get("original_text", ""))) 
    en_orig.pack(pady=8)

    def start_edit_process():
        try:
            num = en_num.get().strip()
            amt_str = en_amt.get().strip()
            orig = en_orig.get().strip()

            if len(num) != 8 or not num.isdigit():
                status.config(text="❌ 號碼須為8位數字")
                return
            amt = float(amt_str)
            if amt < 0:
                status.config(text="❌ 金額不能為負數")
                return

            old_orig = str(old_row.get("original_text", ""))
            
            if orig == old_orig:
                final_save_edit(idx, num, amt, old_row["類別"], orig)
            else:
                res_cat = ai_classify(orig)
                show_edit_confirm(idx, num, amt, res_cat, orig)
        except ValueError:
            status.config(text="❌ 金額請輸入數字")

    create_button(frame, "儲存修改", start_edit_process).pack(pady=15)
    create_button(frame, "取消返回", edit_specific).pack()

# ----------------------------
# 輔助函式
# ----------------------------
def show_edit_confirm(idx, num, amt, cat, orig):
    clear_frame()
    if cat is None:
        tk.Label(frame, text="AI無法判斷,請手動選擇類別", 
                 font=(FONT_MAIN, 11), fg="#FAB1A0", bg=BG_COLOR).pack(pady=15)
        grid_edit = tk.Frame(frame, bg=BG_COLOR)
        grid_edit.pack()
        for i, c in enumerate(["食","衣","住","行","育","樂"]):
            btn = create_button(grid_edit, c, lambda c=c: final_save_edit(idx, num, amt, c, orig), width=8, height=1)
            btn.grid(row=i//2, column=i%2, padx=5, pady=5)
    else:
        tk.Label(frame, text="AI 修改分析結果", font=(FONT_MAIN, 14, "bold"), fg="white", bg=BG_COLOR).pack(pady=15)
        tk.Label(frame, text=f"判斷類別為：{cat}", font=(FONT_MAIN, 12), fg="#3498DB", bg=BG_COLOR).pack(pady=15)
        
        btn_box = tk.Frame(frame, bg=BG_COLOR)
        btn_box.pack(pady=10)
        
        def manual():
            clear_frame()
            tk.Label(frame, text="請選擇正確類別", font=(FONT_MAIN, 11), fg="white", bg=BG_COLOR).pack(pady=10)
            grid_m = tk.Frame(frame, bg=BG_COLOR)
            grid_m.pack()
            for i, c in enumerate(["食","衣","住","行","育","樂"]):
                btn = create_button(grid_m, c, lambda c=c: final_save_edit(idx, num, amt, c, orig), width=8, height=1)
                btn.grid(row=i//2, column=i%2, padx=5, pady=5)
        
        create_button(btn_box, "手動選擇", manual, width=12).grid(row=0, column=0, padx=10)
        create_button(btn_box, "確認正確", lambda: final_save_edit(idx, num, amt, cat, orig), width=12).grid(row=0, column=1, padx=10)

def final_save_edit(idx, num, amt, cat, orig):
    global invoice_data, learning_db 
    
    if orig.strip():
        learning_db[orig] = cat  
        save_learning()          
    
    invoice_data["號碼"] = invoice_data["號碼"].astype(str)
    invoice_data.at[idx, "號碼"] = str(num)
    invoice_data.at[idx, "金額"] = amt
    invoice_data.at[idx, "類別"] = cat
    invoice_data.at[idx, "original_text"] = orig
    
    invoice_data.to_csv(invoice_path, index=False, encoding="utf-8-sig")
    
    edit_specific() 
    status.config(text=f"✅ 修改成功！AI 已同步學習「{orig}」", fg="#55E6C1")

# ----------------------------
if __name__ == "__main__":
    show_main()
    root.mainloop()