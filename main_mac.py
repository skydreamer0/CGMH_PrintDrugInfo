import wx
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
from io import BytesIO
import reportlab.lib.utils
from reportlab.lib.utils import ImageReader
from datetime import datetime, timedelta
import sys
import csv
from io import StringIO
import subprocess
import threading
import PyPDF2

# 註冊中文字體 (Mac 內建: 黑體)
pdfmetrics.registerFont(TTFont('msjh', '/System/Library/Fonts/STHeiti Light.ttc'))
pdfmetrics.registerFont(TTFont('Bold', '/System/Library/Fonts/STHeiti Medium.ttc'))

class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        super(MyFrame, self).__init__(parent, id, title, size=(450, 450))
        global global_font, special_font, special_font2
        
        # 使用 Mac 的字型名稱
        global_font = wx.Font(15, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'Heiti TC')
        special_font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'Heiti TC')
        special_font2 = wx.Font(20, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, u'Heiti TC')
        
        self.SetFont(global_font)
        self.is_query_done = False

        # 設置窗口圖示 (如果找不到不報錯)
        try:
            icon = wx.Icon("CGMH.ico", wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)
        except:
            pass
            
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
        self.print_quantity_combo = wx.ComboBox(panel, size=(100, -1), choices=self.print_quantity_choices,style=wx.CB_READONLY) 
        self.print_quantity_combo.SetSelection(0)

        self.print_button = wx.Button(panel, label="測試列印(Mac預覽)", size=(150, 40))

        print_sizer.Add(self.print_quantity_label, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)
        print_sizer.Add(self.print_quantity_combo, 1, wx.EXPAND | wx.ALL, 5)
        print_sizer.Add(self.print_button, 0, wx.ALIGN_BOTTOM | wx.ALL, 5)

        main_sizer.Add(print_sizer, 0, wx.EXPAND | wx.ALL, 10)

        self.print_button.Bind(wx.EVT_BUTTON, self.on_print)
        self.confirm_button.Bind(wx.EVT_BUTTON, self.on_confirm)
        self.query_input.Bind(wx.EVT_TEXT_ENTER, self.on_confirm)

        self.statusbar.SetFont(special_font)
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
        version_info = "                version.1.2 (Mac修復版)"

        self.statusbar.SetStatusText(author_info, 0)
        self.statusbar.SetStatusText(field_info, 1)
        self.statusbar.SetStatusText(version_info, 2)


    def on_confirm(self, event):
        if not self.is_query_done:
            user_input = self.query_input.GetValue().strip().upper()
            self.last_user_input = user_input
            self.search_and_display(user_input)
            self.is_confirm_button_pressed = True
            self.is_print_button_pressed = False

    def search_and_display(self, user_input):
        result_text = ""
        user_input = str(user_input).upper()

        if not os.path.exists(self.file1) or not os.path.exists(self.file2):
            self.result_textctrl.SetValue("找不到資料檔案 Adgn.txt 或 drugdata.csv")
            return

        # Mac 環境中，處理 Big5/CP950 的編碼問題
        try:
            with open(self.file1, 'r', encoding='cp950', errors='ignore') as f:
                lines = f.readlines()
        except:
            with open(self.file1, 'r', encoding='big5', errors='ignore') as f:
                lines = f.readlines()

        corrected_lines = []
        for line in lines:
            fields = line.split(';')
            corrected_line = ';'.join(fields[:117])
            corrected_lines.append(corrected_line)

        corrected_content = StringIO('\n'.join(corrected_lines))
        df1 = pd.read_csv(corrected_content, delimiter=';')

        try:
            df1 = df1.dropna(subset=['料位號'])
            df1['料位號'] = df1['料位號'].fillna('').astype(str)
            exact_match_result = df1[df1['料位號'] == user_input]
        except KeyError:
            exact_match_result = pd.DataFrame()

        if not exact_match_result.empty:
            料位號 = exact_match_result.iloc[0]['料位號']
            self.write_to_csv(料位號)
            藥品編號 = exact_match_result.iloc[0]['藥品編號']

            try:
                with open(self.file2, 'r', encoding='cp950', errors='ignore') as f:
                    lines = f.readlines()
            except:
                with open(self.file2, 'r', encoding='big5', errors='ignore') as f:
                    lines = f.readlines()

            corrected_lines = []
            for line in lines:
                fields = line.split(',')
                corrected_line = ','.join(fields[:11])
                corrected_lines.append(corrected_line)

            corrected_content = StringIO('\n'.join(corrected_lines))
            df2 = pd.read_csv(corrected_content)

            result2 = df2[df2['藥品編號'] == 藥品編號]
            columns_to_display = ['顏色', '形狀', '劑型', '中文加強描述']
            result2 = result2.fillna('')
            
            try:
                result2_text = result2.to_string(header=False, index=False, col_space=1, justify='left',
                                                 columns=columns_to_display)
            except:
                result2_text = result2.to_string(header=False, index=False, col_space=1, justify='left')

            藥品名稱 = exact_match_result.iloc[0].get('藥品名稱', '')
            條碼 = exact_match_result.iloc[0].get('條碼', '')
            
            qr_image_data = self.generate_qrcode(條碼)
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
        data = str(data)  # 強制轉為字串避免無法編碼
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
        c = canvas.Canvas(os.path.join(os.path.dirname(os.path.abspath(__file__)), "output.pdf"), pagesize=custom_page_size_mm)
        width, height = custom_page_size_mm

        today = datetime.today()
        expiration_date = today + timedelta(days=180)
        formatted_today = today.strftime('%Y-%m-%d')
        formatted_expiration_date = expiration_date.strftime('%Y-%m-%d')

        c.setFont("Bold", 12, leading=8)
        formatted_text = f"{料位號}                  {藥品編號}"
        c.drawString(1, height - 11, formatted_text)

        lines_drug_name = [f"藥名:{藥品名稱}"]
        lines_description = [f"外觀:{外觀描述}"]
        width_limit = width - 10

        wrapped_lines_drug_name = self._wrap_text(lines_drug_name, c, width_limit, "msjh", 9)
        wrapped_lines_description = self._wrap_text(lines_description, c, 99, "msjh", 9)
        wrapped_lines = wrapped_lines_drug_name + wrapped_lines_description

        c.setFont("msjh", 8, leading=99)
        text = c.beginText(1, height - 19)
        text.setFont("msjh", 8, leading=9)

        for line in wrapped_lines:
            text.textLine(line)

        c.drawText(text)
        c.drawImage(qr_image, 99, 2.83, width=40, height=40)

        c.setFont("msjh", 7, leading=7)
        c.drawString(1, height - 70, f"調配日期: {formatted_today}")
        c.drawString(1, height - 80, f"保存期限: {formatted_expiration_date}")
        c.save()

    def _wrap_text(self, lines, canvas_obj, width_limit, font_name, font_size):
        wrapped_lines = []
        for original_line in lines:
            line = str(original_line)
            while line:
                char_idx = 1
                while char_idx <= len(line) and canvas_obj.stringWidth(line[:char_idx], font_name,
                                                                       font_size) <= width_limit:
                    char_idx += 1
                wrapped_lines.append(line[:char_idx - 1].strip())
                line = line[char_idx - 1:]
        return wrapped_lines

    def write_to_csv(self, 料位號):
        current_directory = os.path.dirname(os.path.abspath(__file__))
        csv_file_path = os.path.join(current_directory, 'SearchRecord.csv')

        current_date = datetime.now()
        formatted_date = current_date.strftime("%Y-%m-%d")
        formatted_time = current_date.strftime("%H:%M:%S")

        selected_index = self.print_quantity_combo.GetSelection()
        print_count_str = self.print_quantity_choices[selected_index]
        print_count = int(print_count_str)

        列印調用 = "Y" if self.is_print_button_pressed else "N"

        if os.path.exists(csv_file_path):
            df = pd.read_csv(csv_file_path)
        else:
            df = pd.DataFrame(columns=['日期', '時間', '料位號',  '列印調用', '列印次數'])

        new_row = {'日期': formatted_date, '時間': formatted_time, '料位號': 料位號, '列印調用': 列印調用,'列印次數': print_count}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_csv(csv_file_path, index=False, encoding='utf-8')

    def on_print(self, event):
        print_thread = threading.Thread(target=self.print_in_background)
        print_thread.start()
        selected_index = self.print_quantity_combo.GetSelection()
        print_count_str = self.print_quantity_choices[selected_index]
        self.is_print_button_pressed = True
        print_count = int(print_count_str) if selected_index != -1 else 0

    def print_in_background(self):
        current_directory = os.path.dirname(os.path.abspath(__file__))
        pdf_filename = "output.pdf"
        pdf_path = os.path.join(current_directory, pdf_filename)

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

                    temp_pdf_path = os.path.join(current_directory, f'temp_{_}.pdf')
                    with open(temp_pdf_path, 'wb') as temp_pdf:
                        pdf_writer.write(temp_pdf)

                # 使用 macOS 的 open 指令打開 PDF (模擬列印完成)
                subprocess.call(['open', temp_pdf_path])

            wx.CallAfter(self.search_and_display, self.last_user_input)

        except Exception as e:
            error_message = f"列印預覽時發生錯誤: {e}"
            print(error_message)
            wx.CallAfter(wx.MessageBox, error_message, "列印錯誤", wx.ICON_ERROR)

if __name__ == '__main__':
    app = wx.App(False)
    frame = MyFrame(None, wx.ID_ANY, "林口長庚SATO標籤機藥品標籤列印 (Mac版)")
    frame.Show(True)
    app.MainLoop()
