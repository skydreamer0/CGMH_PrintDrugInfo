import wx
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
from io import BytesIO
from reportlab.lib.utils import ImageReader
from datetime import datetime, timedelta
import sys
import csv
from io import StringIO
import subprocess
import threading
import PyPDF2
import win32api
import win32print

# 註冊中文字體
pdfmetrics.registerFont(TTFont('msjh', 'msjh.ttc'))
pdfmetrics.registerFont(TTFont('Bold', 'msjhbd.ttc'))
class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        super(MyFrame, self).__init__(parent, id, title, size=(400, 400))
        global global_font, special_font, special_font2
        global_font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'微軟正黑體')
        special_font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'微軟正黑體')
        special_font2 = wx.Font(20, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'微軟正黑體')
        self.SetFont(global_font)
        self.is_query_done = False  # 新增這行來追蹤查詢是否完成

        # 設置窗口圖示
        icon = wx.Icon("CGMH.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)
        self.statusbar = self.CreateStatusBar(3)

        self.SetMinSize((400, 400))

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        input_sizer = wx.BoxSizer(wx.HORIZONTAL)
        panel = wx.Panel(self)
        self.query_label = wx.StaticText(panel, label="請輸入料位號")
        self.query_input = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER, size=(-1, 40))
        self.confirm_button = wx.Button(panel, label="確認", size=(100, 40))

        input_sizer.Add(self.query_label, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)
        input_sizer.Add(self.query_input, 1, wx.EXPAND | wx.ALL, 5)
        input_sizer.Add(self.confirm_button, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)

        self.result_label = wx.StaticText(panel, label=" 查找結果：")
        self.result_textctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY)

        main_sizer.Add(input_sizer, 0, wx.EXPAND | wx.ALL, 10)
        main_sizer.Add(self.result_label, 0, wx.ALL, 10)
        main_sizer.Add(self.result_textctrl, 1, wx.EXPAND | wx.ALL, 10)

        # 建立一個新的水平Sizer來容納列印相關的控件
        print_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.print_quantity_label = wx.StaticText(panel, label="選擇列印數量")
        self.print_quantity_choices = ['1',  '5', '10']
        self.print_quantity_combo = wx.ComboBox(panel, size=(100, -1), choices=self.print_quantity_choices,style=wx.CB_READONLY)  # 修改size
        self.print_quantity_combo.SetSelection(0)

        self.print_button = wx.Button(panel, label="列印", size=(100, 40))

        # 將列印數量標籤、下拉選單和列印按鈕添加到新的Sizer中
        print_sizer.Add(self.print_quantity_label, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)
        print_sizer.Add(self.print_quantity_combo, 1, wx.EXPAND | wx.ALL, 5)
        print_sizer.Add(self.print_button, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)

        # 新增新的Sizer到主Sizer
        main_sizer.Add(print_sizer, 0, wx.EXPAND | wx.ALL, 10)

        self.print_button.Bind(wx.EVT_BUTTON, self.on_print)
        self.confirm_button.Bind(wx.EVT_BUTTON, self.on_confirm)
        self.query_input.Bind(wx.EVT_TEXT_ENTER, self.on_confirm)

        # 設置StatusBar的特殊字體
        self.statusbar.SetFont(special_font)
        # 設定ComboBox的特殊字體
        self.print_quantity_combo.SetFont(special_font2)
        panel.SetSizer(main_sizer)

        if getattr(sys, 'frozen', False):
            current_dir = sys._MEIPASS
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))

        self.file1 = os.path.join(current_dir, 'Adgn.txt')
        self.file2 = os.path.join(current_dir, 'drugdata.csv')
        self.set_statusbar_info()
        self.is_confirm_button_pressed = False
        self.is_print_button_pressed = False
        panel.SetSizerAndFit(main_sizer)

    def set_statusbar_info(self):
        author_info = "作者：蕭兆軒 "
        field_info = "查詢快捷鍵為Enter"
        version_info = "                version.1.2"

        self.statusbar.SetStatusText(author_info, 0)
        self.statusbar.SetStatusText(field_info, 1)
        self.statusbar.SetStatusText(version_info, 2)


    def on_confirm(self, event):
        if not self.is_query_done:
            user_input = self.query_input.GetValue().strip().upper()
            self.last_user_input = user_input  # 儲存最後一次有效查詢的 user_input
            self.search_and_display(user_input)
            self.is_confirm_button_pressed = True
            self.is_print_button_pressed = False

    def search_and_display(self, user_input):
        result_text = ""  # 初始化 result_text 為空字符串
        user_input = str(user_input).upper()  # 強制轉換為字符串，然後轉換為大寫

        # 手動讀取和處理 self.file1
        with open(self.file1, 'r', encoding='ANSI') as f:
            lines = f.readlines()

        corrected_lines = []
        for line in lines:
            fields = line.split(';')
            corrected_line = ';'.join(fields[:117])  # 只保留前 117 個字段
            corrected_lines.append(corrected_line)

        corrected_content = StringIO('\n'.join(corrected_lines))
        df1 = pd.read_csv(corrected_content, delimiter=';', encoding='ANSI')

        df1 = df1.dropna(subset=['料位號'])
        df1['料位號'] = df1['料位號'].fillna('').astype(str)

        # 查找料位號完全相符的資訊
        exact_match_result = df1[df1['料位號'] == user_input]
        if not exact_match_result.empty:
            料位號 = exact_match_result.iloc[0]['料位號']
            self.write_to_csv(料位號)
            藥品編號 = exact_match_result.iloc[0]['藥品編號']

            with open(self.file2, 'r', encoding='ANSI') as f:
                lines = f.readlines()

            # 只保留每一行的前11個字段
            corrected_lines = []
            for line in lines:
                fields = line.split(',')
                corrected_line = ','.join(fields[:11])
                corrected_lines.append(corrected_line)

            corrected_content = StringIO('\n'.join(corrected_lines))
            df2 = pd.read_csv(corrected_content, encoding='ANSI')

            result2 = df2[df2['藥品編號'] == 藥品編號]

            columns_to_display = ['顏色', '形狀', '劑型', '中文加強描述']

            # 將外觀部分的"na"替換為空字串
            result2 = result2.fillna('')

            result2_text = result2.to_string(header=False, index=False, col_space=1, justify='left',
                                             columns=columns_to_display)

            藥品名稱 = exact_match_result.iloc[0]['藥品名稱']
            條碼 = exact_match_result.iloc[0]['條碼']
            # 使用條碼資料生成QR碼
            qr_image_data = self.generate_qrcode(條碼)
            # 生成PDF
            self.generate_pdf(料位號, 藥品編號, 藥品名稱, result2_text, qr_image_data)
            result_text = (f"料位號: {料位號}\n"
                           f"藥品編號: {藥品編號}\n"
                           f"藥品名稱: {藥品名稱}\n"
                           f"外觀：\n{result2_text}\n"
                           f"條碼: {條碼}")
        else:
            result_text = "無查詢結果"

        self.result_textctrl.SetValue(result_text)

    def generate_qrcode(self, data):
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=1,
            border=0,
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill='black', back_color='white')
        buffer = BytesIO()
        img.save(buffer, 'PNG')
        buffer.seek(0)
        return ImageReader(buffer)

    def generate_pdf(self, 料位號, 藥品編號, 藥品名稱, 外觀描述, qr_image):
        custom_page_size_mm = (140, 84)
        c = canvas.Canvas("output.pdf", pagesize=custom_page_size_mm)
        width, height = custom_page_size_mm

        # 獲取今天的日期和六個月後的日期
        today = datetime.today()
        expiration_date = today + timedelta(days=180)
        formatted_today = today.strftime('%Y-%m-%d')
        formatted_expiration_date = expiration_date.strftime('%Y-%m-%d')

        # 設定粗體字體 "msjh-Bold"間隔（減少行距）
        c.setFont("Bold", 12, leading=8)
        formatted_text = f"{料位號}                  {藥品編號}"
        c.drawString(1, height - 11, formatted_text)

        # Handle text wrapping
        lines_drug_name = [f"藥名:{藥品名稱}"]
        lines_description = [f"外觀:{外觀描述}"]
        width_limit = width - 10

        wrapped_lines_drug_name = self._wrap_text(lines_drug_name, c, width_limit, "msjh", 9)
        wrapped_lines_description = self._wrap_text(lines_description, c, 99, "msjh", 9)
        wrapped_lines = wrapped_lines_drug_name + wrapped_lines_description

        # Output the wrapped text to PDF
        c.setFont("msjh", 8, leading=99)
        text = c.beginText(1, height - 19)
        text.setFont("msjh", 8, leading=9)


        for line in wrapped_lines:
            text.textLine(line)

        c.drawText(text)

        # Draw QR code
        c.drawImage(qr_image, 99, 2.83, width=40, height=40)

        # Add preparation date and expiration date
        c.setFont("msjh", 7, leading=7)
        c.drawString(1, height - 70, f"調配日期: {formatted_today}")
        c.drawString(1, height - 80, f"保存期限: {formatted_expiration_date}")

        c.save()

    def _wrap_text(self, lines, canvas_obj, width_limit, font_name, font_size):
        wrapped_lines = []
        for original_line in lines:
            line = original_line
            while line:
                char_idx = 1
                while char_idx <= len(line) and canvas_obj.stringWidth(line[:char_idx], font_name,
                                                                       font_size) <= width_limit:
                    char_idx += 1
                wrapped_lines.append(line[:char_idx - 1].strip())
                line = line[char_idx - 1:]
        return wrapped_lines

    def write_to_csv(self, 料位號):
        csv_file_path = 'SearchRecord.csv'  # 設定新CSV文件的路徑

        current_date = datetime.now()
        formatted_date = current_date.strftime("%Y-%m-%d")
        formatted_time = current_date.strftime("%H:%M:%S")

        # 獲取列印相關的信息
        selected_index = self.print_quantity_combo.GetSelection()
        print_count_str = self.print_quantity_choices[selected_index]
        print_count = int(print_count_str)

        # 判斷是否按下列印按鈕
        列印調用 = "Y" if self.is_print_button_pressed else "N"

        # 開啟或創建一個CSV檔案並寫入資料
        if os.path.exists(csv_file_path):
            df = pd.read_csv(csv_file_path)
        else:
            df = pd.DataFrame(
                columns=['日期', '時間', '料位號',  '列印調用', '列印次數'])  # 初始化一個空的 DataFrame，並添加新的列

        new_row = {'日期': formatted_date, '時間': formatted_time, '料位號': 料位號, '列印調用': 列印調用,'列印次數': print_count}
        # 使用concat方法將新行轉換為DataFrame，然後連接到現有的DataFrame
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        df.to_csv(csv_file_path, index=False, encoding='utf-8')


    def on_print(self, event):
        # 創建一個新的線程來執行 print_in_background 函數
        print_thread = threading.Thread(target=self.print_in_background)
        print_thread.start()
        selected_index = self.print_quantity_combo.GetSelection()
        print_count_str = self.print_quantity_choices[selected_index]
        self.is_print_button_pressed = True
        print_count = int(print_count_str) if selected_index != -1 else 0  # 如果未按下列印按鈕，列印次數為0

    def print_in_background(self):
        current_directory = os.path.dirname(os.path.abspath(__file__))
        pdf_filename = "output.pdf"
        pdf_path = os.path.join(os.getcwd(), pdf_filename)

        selected_index = self.print_quantity_combo.GetSelection()
        print_count_str = self.print_quantity_choices[selected_index]
        print_count = int(print_count_str)

        try:
            for _ in range(print_count):
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    pdf_writer = PyPDF2.PdfWriter()

                    for page_num in range(len(pdf_reader.pages)):
                        pdf_writer.add_page(pdf_reader.pages[page_num])

                    temp_pdf_path = os.path.join(os.getcwd(), 'temp.pdf')
                    with open(temp_pdf_path, 'wb') as temp_pdf:
                        pdf_writer.write(temp_pdf)

                # 使用 win32api 列印 PDF
                win32api.ShellExecute(0, "print", temp_pdf_path, None, ".", 0)

            # 使用 wx.CallAfter 來確保 GUI 更新是線程安全的
            wx.CallAfter(self.search_and_display, self.last_user_input)  # 使用最後一次有效的 user_input

        except Exception as e:
            error_message = f"列印時發生錯誤: {e}"
            print(error_message)
            wx.CallAfter(wx.MessageBox, error_message, "列印錯誤", wx.ICON_ERROR)


if __name__ == '__main__':
    app = wx.App(False)
    frame = MyFrame(None, wx.ID_ANY, "林口長庚SATO標籤機藥品標籤列印")
    frame.Show(True)
    app.MainLoop()