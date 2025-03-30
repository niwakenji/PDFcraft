# PDFcraft.py
# Copyright (c) 2025 Kenji Niwa / koromokkuru lab.

# This work is licensed under the Creative Commons Attribution 4.0 International License.
# To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/

# このコードは クリエイティブ・コモンズ 表示 4.0 国際ライセンス（CC BY 4.0）に基づき提供されています。
# 商用利用、改変、再配布は自由ですが、出典の表示（著作者名とライセンスURL）をお願いします。
# クレジット例: "PDFcraft by Kenji Niwa is licensed under CC BY 4.0"

#v1.0 2025/3/30

import os
import sys
import tkinter as tk

def show_loading_message():
    loading = tk.Tk()
    loading.title("起動中...")
    loading.geometry("300x100")
    label = tk.Label(loading, text="ツールを初期化中です...\nしばらくお待ちください", font=("Arial", 10))
    label.pack(expand=True)
    loading.update()
    return loading

_loading_win = show_loading_message()
import shutil
from PIL import Image
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import simpledialog, filedialog
import win32com.client
import pythoncom
import pyautogui
import random
import ast
import re
from datetime import datetime, timedelta
import time
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from pdf2docx import Converter  # ← PDF→Word 変換ライブラリ
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5


#c ユーザーフォーム ★
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32com.client

class CommandMaker:
    def __init__(self, pdfcraft):
        self.txt_message = "\nここにファイルをドロップしてください\n"
        self.root = TkinterDnD.Tk()
        self.drop_label = None
        self.pdfcraft = pdfcraft

    def run(self):
        self._setup_window()
        self._setup_usage_table()
        self._setup_drop_label()
        self.root.mainloop()

    def _setup_window(self):
        self.root.title("ツールの使い方")
        self.root.resizable(False, False)
        self.root.geometry("331x215")
        self.root.bind("<Double-Button-1>", self.on_double_click)
        self.root.bind("<Button-1>", self.on_single_click)

    def _setup_usage_table(self):
        usage_text = [
            ("ドロップ内容", "動作"),
            ("なし(=dbClick)", "パワポ 透かし総ﾍﾟｰｼﾞPDF化"),
            ("PDF 1つ", "分割 /透かし /pw付 /Word変換"),
            ("PDF 複数", "結合 /置換 /挿入"),
            ("画像複数", "PDF化"),
        ]

        for col, header in enumerate(usage_text[0]):
            label = tk.Label(self.root, text=header, font=("Arial", 11, "bold"),
                             borderwidth=2, relief="ridge", padx=10, pady=5)
            label.grid(row=0, column=col, sticky="nsew")

        for row, (col1, col2) in enumerate(usage_text[1:], start=1):
            label1 = tk.Label(self.root, text=col1, font=("Arial", 10),
                              borderwidth=1, relief="solid", padx=10, pady=4)
            label2 = tk.Label(self.root, text=col2, font=("Arial", 10),
                              borderwidth=1, relief="solid", padx=10, pady=4)
            label1.grid(row=row, column=0, sticky="nsew")
            label2.grid(row=row, column=1, sticky="nsew")

    def _setup_drop_label(self):
        self.drop_label = tk.Label(
            self.root,
            text=self.txt_message,
            font=("Arial", 10),
            fg="gray",
            wraplength=300,
            justify="center"
        )

        self.drop_label.grid(row=5, column=0, columnspan=2, pady=(8, 0))
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.on_drop)

    def update_message(self, msg):
        if self.drop_label:
            self.drop_label.config(text = msg)

    def on_drop(self, event): #★
        input_paths = list(self.root.tk.splitlist(event.data)) #list() 重要。list()がないと、タプル()になる。
        
        # CommandMaker
        result = self._Lv0selector(input_paths) #ドロップされたアイルの種類によって分岐させ、実行し、message[]を返す
        commands_file_path = os.path.dirname(os.path.abspath(sys.argv[0])) + "\commands.txt"
        if result:
            self._SaveResultToTxt(result, commands_file_path) # resultの内容を保存
            pdcraft_result = self.pdfcraft.run(commands_file_path) # コマンドファイルの実行
            self.update_message(pdcraft_result["message"]) # pdcraftの結果を表示
        else:
            self.update_message("\nキャンセルされました") #
    
    def _SaveResultToTxt(self, result, commands_file_path): # resultの内容を保存する
            commands_list = result["commandlines"]
            with open(commands_file_path, "w", encoding="utf-8") as file: #ファイルへ書き込み
                commands_list = commands_list.replace("'", '"') # シングルをダブルコーテーションに変換
                print(f"コマンドを出力しました：{commands_file_path}")
                file.write(commands_list)  # 各要素を改行で結合して書き込む
            
 
    def on_double_click(self, event):
        input_ppt_instance = self._get_powerpoint_instance() # self._get_powerpoint_address()
         
        if input_ppt_instance:
            input_ppt_path = input_ppt_instance.FullName
            input_ppt_name = os.path.basename(input_ppt_path)
            last_page = self._ppt_get_last_page(input_ppt_instance)  # last_pageを取得
            result = self._create_commandslines_copy_pptwatermark(input_ppt_path, last_page)

            # 出力⇒実行
            commands_list = result["commandlines"]
            commands_file_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "commands.txt") # アドレスの区切りを自動処理
            with open(commands_file_path, "w", encoding="utf-8") as file: #ファイルへ書き込み
                commands_list = commands_list.replace("'", '"') # シングルをダブルコーテーションに変換
                print(f"コマンドを出力しました：{commands_file_path}")
                file.write(commands_list)  # 各要素を改行で結合して書き込む
                
            pdcraft_result = self.pdfcraft.run(commands_file_path) # コマンドファイルの実行
            self.update_message(pdcraft_result["message"]) # pdcraftの結果を表示
            #self.update_message(result["message"]) # 事前に得ていたメッセージを表示

            self.update_message(result["message"])
        else:
            self.update_message("パワーポイントを開いた状態でダブルクリック\n"
                                "1ページ目の「透かし」「総ページ」と名付\n"
                                "けた要素を全ページに複製しPDF保存します")

    def on_single_click(self, event):
        self.update_message("PDFを編集するPythonアプリです\nこのウィンドウへの\nファイルドロップで動作します")

    # --- フォルダ内ファイルの拡張子を調べる ---
    def _analyze_input_paths(self, input_paths):
        """
        ファイルパス一覧を受け取り、拡張子が統一されているか調べる。

        - 統一されていれば → (件数, 拡張子)
        - 混在 or 拡張子なし → (件数, -1)
        """
        found_ext = None
        for path in input_paths:
            ext = os.path.splitext(path)[1].lower()
            if not ext:
                return len(input_paths), -1  # 拡張子なしがあれば即NG

            if found_ext is None:
                found_ext = ext  # 最初の拡張子を記録
            elif ext != found_ext:
                return len(input_paths), -1  # 異なる拡張子があればNG

        return len(input_paths), found_ext if input_paths else (0, -1)


    def _check_folder_extension(self, folder_path):
        """
        指定フォルダ内のファイル（自分自身とcommands.txtとディレクトリを除く）を返す
        拡張子が統一されていればその拡張子、混在していれば -1 を返す
        """
        print (f"_check_folder_extension 引数：{folder_path}")
        exe_name = os.path.basename(sys.argv[0]).lower()
        input_paths = []

        for f in os.listdir(folder_path):
            full_path = os.path.join(folder_path, f)

            if not os.path.isfile(full_path):  # フォルダは除外
                continue
            if f.lower() == exe_name or f.lower() == "commands.txt" :  # 自分自身は除外
                continue

            input_paths.append(full_path)  # 有効ファイルを追加

        if not input_paths:
            return 0, -1, []  # 空なら即NG

        count, ext = self._analyze_input_paths(input_paths)
        return count, ext, input_paths  # 常にファイル一覧は返す

    def _get_powerpoint_address(self):
        #開いているPowerPointプレゼンテーションのパスを取得する関数　複数ある場合は選択
        pythoncom.CoInitialize()

        try:
            pptApp = win32com.client.Dispatch("PowerPoint.Application")

            presentations_count = pptApp.Presentations.Count

            if presentations_count == 0:
                messagebox.showerror("エラー", "開いているPowerPointプレゼンテーションがありません。")
                return None
            elif presentations_count == 1:
                return pptApp.Presentations.Item(1).FullName
            else:
                ppt_list = [pptApp.Presentations[i].Name for i in range(presentations_count)]
                options = "\n".join(f"{i + 1}: {name}" for i, name in enumerate(ppt_list))
                choice = simpledialog.askinteger("プレゼンテーション選択", f"開いているプレゼンテーションを番号で選んでください:\n{options}", minvalue=1, maxvalue=presentations_count)
                if choice is not None:
                    return pptApp.Presentations.Item(choice).FullName
                else:
                    print("ユーザーが選択をキャンセルしました。")
                    return None

        except Exception as e:
            # messagebox.showerror("エラー", f"PowerPointを起動できませんでした。\n\n{e}")
            return None

    def _ppt_get_last_page(self, ppt_instance):
            #PDF化するページ数を取得する関数"""
            total_pages = ppt_instance.Slides.Count
            last_page = simpledialog.askinteger("パワポ透かし追加＆PDF化", f"{ppt_instance.Name} の何ページ目までをPDFにしますか？（1～{total_pages}）", minvalue=1, maxvalue=total_pages)
            return last_page


    def _get_powerpoint_instance(self):
        # 開いているPowerPointプレゼンテーションのインスタンスを取得する関数。複数ある場合は選択。
        pythoncom.CoInitialize()

        try:
            pptApp = win32com.client.Dispatch("PowerPoint.Application")
            presentations_count = pptApp.Presentations.Count

            if presentations_count == 0:
                messagebox.showerror("エラー", "開いているPowerPointプレゼンテーションがありません。")
                return None
            elif presentations_count == 1:
                return pptApp.Presentations.Item(1)
            else:
                ppt_list = [pptApp.Presentations[i].Name for i in range(presentations_count)]
                options = "\n".join(f"{i + 1}: {name}" for i, name in enumerate(ppt_list))
                choice = simpledialog.askinteger(
                    "プレゼンテーション選択",
                    f"開いているプレゼンテーションを番号で選んでください:\n{options}",
                    minvalue=1,
                    maxvalue=presentations_count
                )
                if choice is not None:
                    return pptApp.Presentations.Item(choice)
                else:
                    print("ユーザーが選択をキャンセルしました。")
                    return None

        except Exception as e:
            # messagebox.showerror("エラー", f"PowerPointを起動できませんでした。\n\n{e}")
            return None


    def _create_commandslines_copy_pptwatermark(self, input_ppt_path, last_page): #PPT透かしコピーのコマンドラインを作成
        input_ppt_path = input_ppt_path.replace("\\", "\\\\")  # ←これ追加！
        output_pdf_path = os.path.dirname(input_ppt_path) + "\\\\water.pdf"
        output_fname = os.path.basename(output_pdf_path) # 出力ファイル名（フォルダなし

        commandlines = f'sukashi("{input_ppt_path}", {last_page},["透かし","総ページ"],["総ページ"], "{output_pdf_path}")'
        print(commandlines)
        return {
            "status": 1,
            "commandlines": commandlines, 
            "message": f"フォルダ{os.path.dirname(input_ppt_path)}に、{output_fname}を保存しました。",
            "log": "",
        }


    def _embed_watermark(self, input_paths): #透かしをつける PDF 1～複数に対応 ★
        
        if isinstance(input_paths, str):
            input_paths = [input_paths]  # リストの1要素にする
        
        watermark = simpledialog.askstring("透かし文字", "PDFに表示する透かし文字を入力してください：")
        if not watermark:
            watermark = "Confidential"
            
        output_paths = [file.replace(".pdf", "_watermark.pdf") for file in input_paths] # 出力ファイル名
        # input_paths =  [file.replace("_透かし", "") for file in output_paths] # 出力ファイル名
        output_fname = [os.path.basename(file) for file in input_paths] # 出力ファイル名（フォルダなし）
        commandlines = '# 透かし文字\n# PDFファイルに透かし文字列を追加し、新しいファイル名で保存します。\n# watermark(["元PDF1.pdf", "元PDF2.pdf", ...], パスワード, ["出力PDF1.pdf", "出力PDF2.pdf", ...])\n'
        return {
            "status": 1,
            "commandlines": commandlines + f'watermark({input_paths}, "{watermark}", {output_paths})', 
            "message":  f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
            "log": "",
        } 
   
   
    def _merge(self, input_paths): #pdfを結合 PDF 2～複数に対応
        input_paths = self._simpledialog_reorder_filepaths(input_paths) #ユーザーの選択による並べ替え　結合順の変更
        
        if input_paths:
            output_path = os.path.dirname(input_paths[0]) + "\merge.pdf"
            output_fname = os.path.basename(output_path) # 出力ファイル名（フォルダなし)
            commandlines = '# 結合\n# 複数のPDFファイルを1つのPDFにまとめます。結合するPDFのファイルパスはリスト形式で記述します。\n# merge(["元PDF1.pdf", "元PDF2.pdf", ...], "出力PDF.pdf")\n'
            return {
                "status": 1,
                "commandlines": commandlines + f'merge({input_paths}, "{output_path}")', 
                "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
                "log": "",
            }
        else:
            return None

    def _attach_password(self, input_paths): #パスワードをつける PDF １～複数に対応
        if isinstance(input_paths, str):
            input_paths = [input_paths]  # リストの1要素にする
    
        password = simpledialog.askstring("パスワード入力", "PDFに設定するパスワードを入力してください：\n空欄でOKを押すと自動生成されます" )    
        message1 = ""
        random_decimal = ""
        if not password:
            # ランダムな10進数の値を生成（0から4294967295まで）
            random_decimal = random.randint(0, 0xFFFFFFFF)  # 0xFFFFFFFFは32ビットの最大値
            # 16進数に変換し、8桁になるようにフォーマット
            password = f"{random_decimal:08X}"  # 大文字の16進数
            password = password.lower() # 小文字
            # password = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
            message1 = f"自動設定パスワード：{password}\n"
            
        output_paths = [file.replace(".pdf", f"_pass_{password}.pdf") for file in input_paths] # 出力ファイル名
        output_fname = [os.path.basename(file) for file in input_paths] # 出力ファイル名（フォルダなし）
        commandlines = '# パスワード付与\n# PDFにパスワードを付加し、新しいPDFとして保存します。\n# password(["元PDF1.pdf", "元PDF2.pdf", ...], パスワード, ["出力PDF1.pdf", "出力PDF2.pdf", ...])\n'
        return {
            "status": 1,
            "commandlines": commandlines + f'password({input_paths}, "{password}", {output_paths})', 
            "message":  f"{message1}フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
            "log": "",
        }

    def _get_page_range_list(self, page_str): # ユーザーのページ番号入力を配列に格納 1 -> [1,1] 2-2 -> [2,2] 1-5 -> [1,5]
        if "-" in page_str:
            start, end = page_str.split("-")
        else:
            start = end = page_str
        return [start.strip(), end.strip()]
    
    def _get_pdf_total_page(self, pdf_path): #PDFファイルのページ数
        reader = PdfReader(pdf_path)
        return len(reader.pages)

    def _Lv0selector(self, input_paths):
        
        #log("1: ドロップファイルの数と拡張子を取得します")
        count1, ext1 = self._analyze_input_paths(input_paths) #ファイルの拡張子とパスを取得（複数）
        #log("2: 起動フォルダ内のファイルの数、拡張子、パスを取得します")
        count2, ext2, input_paths2 = self._check_folder_extension(os.path.dirname(os.path.abspath(sys.argv[0]))) #同一拡張子の数、同一拡張子名、パス全部を取得
        
        #log("ドロップファイルの有無、数に応じて実行内容を分岐させます")
        if ext1 == ".pdf":
            if count1 == 1:
                if count2 >= 1 and ext2 in [".jpg", ".jpeg", ".png", ".bmp"]: #起動フォルダのファイル情報
                    return self._Lv1Selector_PDF1_JPG(input_paths, input_paths2) #画像をPDFに追加
                else:
                    return self._Lv1Selector_PDF1(input_paths) #分割 抽出 ページ削除 透かし　パスワード付与 (Word変換)
            elif count1 == 2:
                return self._Lv1Selector_PDF2(input_paths) #結合　置換 挿入
            elif count1 > 2:
                return self._Lv1Selector_PDF3over(input_paths) #結合
        elif ext1 in [".jpg", ".jpeg", ".png", ".bmp"]:
            return self._Lv1Selector_JPG(input_paths) #画像PDF化
        elif ext1 in [".txt"]:
            return self._Lv1Selector_TXT(input_paths) #テキストファイル（複数可）のコマンドを実行
        else:
            return {
                "status": -1,
                "commandlines": "",
                "message": "非対応のファイルがドロップされました",
                "log": "",
            }


    def _Lv1Selector_PDF1_JPG(self, input_paths, input_paths2): #画像をPDFに追加
        file_name = os.path.basename(input_paths[0])
        image_paths = sorted(input_paths2, key=lambda x: os.path.basename(x).lower())  # ファイル名でソート
        backup_path = add_date_to_filename(input_paths[0]) #末尾に日付付加
        backup_name = os.path.basename(backup_path)
        commandlines = '# 画像をPDFに追加\n# 画像ファイルをPDFに変換してから、元のPDFと結合します。元PDFは日付付きでバックアップされます。\n# add("元PDF.pdf", ["画像1.jpg", "画像2.jpg", ...], "出力PDF.pdf")\n'
        return {
            "status": 1,
            "commandlines": commandlines + f'add("{input_paths[0]}", {image_paths}, "{backup_path}")', 
            "message": f"{file_name}に画像を追加し、元ファイルを{backup_name}に変更しました。",
            "log": "",
        }
    
    
    def _Lv1Selector_PDF1(self, input_paths): #分割 抽出 ページ削除 透かし　パスワード付与 (Word変換)
        pdf_page_range = "1-" + str(self._get_pdf_total_page(input_paths[0]))
        user_input = simpledialog.askstring("処理選択",
                                            "\n実行したい処理を選んでください：\n"
                                            "保存先：ドロップファイルの横\n\n"
                                            "1 : PDFの分割\n"
                                            "2 : PDFの抽出\n"
                                            "3 : PDFに削除\n" 
                                            "4 : PDFに透かし付与\n" 
                                            "5 : PDFにパスワード付与\n" )

        if user_input is None:
            return {
                "status": 0,
                "commandlines": "", 
                "message": "キャンセルされました",
                "log": "",
            }
        elif user_input.strip() == "1": # 分割
            page = simpledialog.askstring("分割ページ", f"分割ページを入力してください：{pdf_page_range}")
            output_path1 = input_paths[0].replace(".pdf", "_part1.pdf")
            output_path2 = input_paths[0].replace(".pdf", "_part2.pdf")
            output_fname1 = os.path.basename(output_path1)
            output_fname2 = os.path.basename(output_path2)
            commandlines = '# 分割\n# 分割ページ番号を指定すると、その前までのページが前半、残りが後半になります。\n# split("元PDF.pdf", 分割ページ番号, "前半出力.pdf", "後半出力.pdf")\n'
            return {
                "status": 1,
                "commandlines": commandlines + f'split("{input_paths[0]}", {page}, "{output_path1}", "{output_path2}")', 
                "message":  f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname1}, {output_fname2}を保存しました。",
                "log": "",
            }
        elif user_input.strip() == "2": # 抽出       
             page = simpledialog.askstring("抽出ページ", "抽出ページを入力してください:", initialvalue = pdf_page_range)
             output_path = input_paths[0].replace(".pdf", "_extract.pdf")
             output_fname = os.path.basename(output_path)
             pages = self._get_page_range_list(page) # ユーザーのページ入力(1, 1-1, 1-3) を ([1,1] [1,1] [1,3])に変換
             commandlines = '# 抽出\n# 指定されたページ範囲のみを抽出して、新しいPDFとして保存します。\n# extract("元PDF.pdf", 開始ページ, 終了ページ, "出力PDF.pdf")\n'
             return {
                "status": 1,
                "commandlines": commandlines + f'extract("{input_paths[0]}", {pages[0]}, {pages[1]}, "{output_path}")', 
                "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
                "log": "",
             }
        elif user_input.strip() == "3": # ページ削除
             page = simpledialog.askstring("削除ページ", "削除ページを入力してください：",  initialvalue = pdf_page_range)
             output_path = input_paths[0].replace(".pdf", "_remove.pdf")
             output_fname = os.path.basename(output_path)
             pages = self._get_page_range_list(page) # ユーザーのページ入力(1, 1-1, 1-3) を ([1,1] [1,1] [1,3])に変換
             commandlines = '# 削除\n# 指定したページ範囲を元PDFから削除し、新しいPDFを作成します。\n# remove("元PDF.pdf", 開始ページ, 終了ページ, "出力PDF.pdf")\n'
             return {
                "status": 1,
                "commandlines": commandlines + f'remove("{input_paths[0]}", {pages[0]}, {pages[1]}, "{output_path}")', 
                "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
                "log": "",
             }
        elif user_input.strip() == "4": # 透かし付与
            return self._embed_watermark(input_paths)
        elif user_input.strip() == "5": # パスワード付与
            return self._attach_password(input_paths)
        else:
            return {
                "status": -1,
                "commandlines": "",
                "message": "エラー 有効な数字を入力してください",
                "log": "",
            }
    
    def _select1_of_2files(self, input_paths): #2つのファイルから1つを選び、input_pathsの順番を入れ替える
        input_paths = list(input_paths)  # タプル対策！
        filelist = self._str_filepaths_to_fielist(input_paths) # ファフィルリストのテキスト
        if filelist:
            base_no = simpledialog.askstring("ベース選択", f"ベースファイルを数字で入力してください(1 or 2)\n{filelist}")
            print ("--------nnnnn-----------")
            print (base_no)
            if base_no == "2": #順番の入れ替え [0]がベースファイル
                input_paths = [input_paths[1], input_paths[0]]
            
            return input_paths
        else:
            return None
    
        
    def _Lv1Selector_PDF2(self, input_paths): #結合　置換 挿入 パスワード付与
        user_input = simpledialog.askstring("処理選択",
                                            "\n実行したい処理を選んでください：\n"
                                            "保存先：ドロップファイルの横\n\n"
                                            "1 : PDFの結合\n"
                                            "2 : PDFの置換\n"
                                            "3 : PDFの挿入\n"
                                            "4 : PDFに透かし付与\n"
                                            "5 : PDFにパスワード付与\n")

        if user_input is None:
            return {
                "status": 0,
                "commandlines": "", 
                "message": "キャンセルされました",
                "log": "",
            }
        elif user_input.strip() == "1": # ファイル結合
            return self._merge(input_paths)
        elif user_input.strip() == "2": # 置換
            input_paths = self._select1_of_2files(input_paths) #２つから1つを選択し、入れ替え
            pdf_page_range = "1-" + str(self._get_pdf_total_page(input_paths[0]))
            if input_paths:
                page = simpledialog.askstring("置換ページ", "置換ページを入力してください：", initialvalue = pdf_page_range)
                pages = self._get_page_range_list(page) # ユーザーのページ入力(1, 1-1, 1-3) を ([1,1] [1,1] [1,3])に変換
                output_path = input_paths[0].replace(".pdf", "_replace.pdf")
                output_fname = os.path.basename(output_path)
                commandlines = '# 置換\n# 指定したページ範囲（開始～終了）を、別のPDFの内容で置き換えます。\n# replace("元PDF.pdf", 開始ページ, 終了ページ, "差替えPDF.pdf", "出力PDF.pdf")\n'
                return {
                    "status": 1,
                    "commandlines": commandlines + f'replace("{input_paths[0]}", {pages[0]}, {pages[1]}, "{input_paths[1]}", "{output_path}")', 
                    "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
                    "log": "",
                }
            else:
                return None
                
        elif user_input.strip() == "3": # 挿入
            input_paths = self._select1_of_2files(input_paths) #２つから1つを選択し、入れ替え
            if input_paths:
                page = simpledialog.askstring("挿入ページ", "指定ページの前に挿入します。ページを入力してください：")                    
                output_path = input_paths[0].replace(".pdf", "_insert.pdf")
                output_fname = os.path.basename(output_path)
                commandlines = '# 挿入\n# PDFファイルを挿入し1つのPDFとして保存します。\n# insert("元PDF.pdf", 挿入ページ, "挿入PDF.pdf", "出力PDF.pdf")\n'
                return {
                    "status": 1,
                    "commandlines": commandlines + f'insert("{input_paths[0]}", {page}, "{input_paths[1]}", "{output_path}")', 
                    "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_fname}を保存しました。",
                    "log": "",
                }
            else:
                return None
                
        elif user_input.strip() == "4": # 透かし付与
            return self._embed_watermark(input_paths)
        elif user_input.strip() == "5": # パスワード付与
            return self._attach_password(input_paths)
        else:
            return {
                "status": -1,
                "commandlines": "",
                "message": "エラー 有効な数字を入力してください",
                "log": "",
            }
        
    def _Lv1Selector_PDF3over(self, input_paths): #結合 パスワード付与 透かし
        from tkinter import Tk, simpledialog, messagebox
        # ユーザー入力
        root = Tk()
        root.withdraw()
        user_input = simpledialog.askstring("処理選択",
                                            "\n        実行したい処理を選んでください："
                                            "\n保存先：ドロップファイルの横\n\n"
                                            "1 : PDFの結合\n"
                                            "2 : 透かし付与\n"
                                            "3 : PDFにパスワード付与\n")

        if user_input is None:
            return {
                "status": 0,
                "commandlines": "", 
                "message": "キャンセルされました",
                "log": "",
            }
        elif user_input.strip() == "1": # ファイル結合
            return self._merge(input_paths)
        elif user_input.strip() == "2": # 透かし付与
            return self._embed_watermark(input_paths)
        elif user_input.strip() == "3": # パスワード付与
            return self._attach_password(input_paths)
        else:
            return {
                "status": -1,
                "commandlines": "",
                "message": "エラー 有効な数字を入力してください",
                "log": "",
            }

    def _Lv1Selector_JPG(self, input_paths): #画像PDF化()
        input_paths = sorted(input_paths, key=lambda x: os.path.basename(x).lower())  # ファイル名でソート
        output_path = os.path.dirname(input_paths[0]) + "\image.pdf"
        output_path = add_date_to_filename(output_path) #末尾に日付付加
        output_name = os.path.basename(output_path)
        commandlines = '# 画像PDF化\n# 複数の画像ファイルを1つのPDFとして保存します。順番はファイル名の昇順です。\n# pdf(["画像1.jpg", "画像2.png", ...], "出力PDF.pdf"\n'
        return {
            "status": 1,
            "commandlines": commandlines + f'pdf({input_paths}, "{output_path}")', 
            "message": f"フォルダ{os.path.dirname(input_paths[0])}に、{output_name}を保存しました。",
            "log": "",
        }

    def _Lv1Selector_TXT(self, input_paths): #複数TXTのコマンドを返す
        # コメントは削除して、配列に入れて返す。TXTは複数OK
        commnand_lines = read_commands_with_substitution(input_paths)
        input_name = [os.path.basename(file) for file in input_paths] # 出力ファイル名（フォルダなし）
        return {
            "status": 1,
            "commandlines": commnand_lines, 
            "message": f'{input_name}を実行しました',
            "log": "",
        }
    
    # --- ユーザー入力ボックス ---
    #input_paths の順番を入れ替える
    #9個以下なら → 123 のような1桁数字で順番指定  10個以上なら → カンマ区切り（例：1,3,5,10）
    def _str_filepaths_to_fielist(self, input_paths): #ファイルパスからファイル名のリストをテキストで返す
        #abc順にならべる
        input_paths = sorted(input_paths, key=lambda x: os.path.basename(x).lower())  # ファイル名でソート
        
        file_list = ""
        for i, path in enumerate(input_paths, start=1):
            file_list += f"{i}: {os.path.basename(path)}\n"

        return file_list     

    def _simpledialog_reorder_filepaths(self, input_paths):
        
        file_list = self._str_filepaths_to_fielist(input_paths) #ファイルパスからファイル名のリストをテキスト
        count_files = len(input_paths)
        input_paths = sorted(input_paths, key=lambda x: os.path.basename(x).lower())  # ファイル名でソート
        
        # 入力形式の説明（10個未満は1桁連続、10個以上はカンマ区切り）
        if count_files < 10:
            example = "例：321"
            hint = "※半角数字で、結合したい順番を入力（連続）"
        else:
            example = "例：1,3,5,10"
            hint = "※カンマ区切りで、結合したい順番を入力"

        prompt = f"結合したいPDFの順を入力してください（部分選択も可能)\n保存先：ドロップファイルの横\n\n{hint}\n{example}\n\n{file_list}"

        order_input = simpledialog.askstring("PDFの結合", prompt)
        print(f"order_input:{order_input}")
        if order_input == "":
            return input_paths #ソート済

        order_input = order_input.strip()

        try:
            if count_files < 10:
                # 1文字ずつ分割（例：312 → [3,1,2]）
                indices = [int(c) - 1 for c in order_input if c.isdigit()]
            else:
                # カンマで分割（例："1,3,10" → [0,2,9]）
                parts = [int(s.strip()) - 1 for s in order_input.split(",")]
                indices = parts

            # 範囲外のインデックスがないかチェック
            if any(i < 0 or i >= count_files for i in indices):
                raise ValueError
        except:
            return -2

        ordered_paths = [input_paths[i] for i in indices]

        return ordered_paths


    
#c Parser            ★
class CommandParser:
    def __init__(self):
        # 
        pass

    def parse(self, command_line: str):
        """
        文字列から関数名と引数を抽出し、Pythonオブジェクトとして返す。
        例: 'merge("merge", "file1.pdf", "file2.pdf", "out.pdf")'
          → ('merge', ['merge', 'file1.pdf', 'file2.pdf', 'out.pdf'])
        """
        raw_args = self.extract_function_and_args(command_line)
        if not raw_args:
            raise ValueError("[CommandParser.parse]Invalid command format")

        func_name = raw_args[0]  # 最初の要素が関数名
        args = [ast.literal_eval(arg) for arg in raw_args[1:]]  # 引数を文字列→Python値に変換
        args.insert(0, func_name) # 引数に、関数名（ユーザー名称）を追加
        
        return func_name, args

    def extract_function_and_args(self, command_str: str):
        """
        "RegisterWordToDictionary(A,[1,3,4,5],C)" のような文字列から
        関数名と引数（配列も含む）を抽出し、FunctionArray に格納する関数
        """
        print(command_str)
        FunctionArray = []

        # 関数名と引数部分を取得
        if "(" not in command_str or ")" not in command_str:
            return FunctionArray

        # "(" で2つに分割し、関数名と引数文字列に分ける
        parts = command_str.split("(", 1)
        FunctionName = parts[0].strip()
        ArgsString = parts[1].rstrip(")").strip()  # 最後の ) を取り除く

        FunctionArray.append(FunctionName)
        
        # 引数部分を1文字ずつ解析し、カンマ区切りで配列に分割（中括弧内のカンマは無視）
        arg = ""
        depth = 0
        for char in ArgsString:
            if char == "[":
                depth += 1
            elif char == "]":
                depth -= 1
            elif char == "," and depth == 0:
                FunctionArray.append(arg.strip())
                arg = ""
                continue
            arg += char

        # 最後に残る引数を追加
        if arg.strip():
            FunctionArray.append(arg.strip())

        return FunctionArray

    def parse_scheduled(self, command_line: str):
        """
        スケジュール付きコマンドを解析。
        例: 'at now + 2min do merge(...)' → (datetimeオブジェクト, "merge(...)")
        """
        match = re.match(r'at (.+?) do (.+)', command_line)
        if not match:
            raise ValueError("[CommandParser.parse_scheduled]Invalid scheduled command syntax")

        time_str, command_part = match.groups()

        # 相対時間指定（now + ●min / ●sec）
        if time_str.strip().startswith("now +"):
            delta_str = time_str.strip()[5:].strip()
            if delta_str.endswith("min"):
                minutes = int(delta_str[:-3].strip())
                scheduled_time = datetime.now() + timedelta(minutes=minutes)
            elif delta_str.endswith("sec"):
                seconds = int(delta_str[:-3].strip())
                scheduled_time = datetime.now() + timedelta(seconds=seconds)
            else:
                raise ValueError(f"Unknown time unit in: {delta_str}")
        else:
            # 絶対時刻（ISO形式）
            scheduled_time = datetime.fromisoformat(time_str.strip())

        return scheduled_time, command_part.strip()
    
#c operator                ★
class CommandOperator:
    def __init__(self, parser, logger=None):
        self.command_map = None
        self.parser = parser
        self.logger = logger  # ← ロガー追加

    def set_command_map (self, command_map):
        self.command_map = command_map
    
    def execute(self, func_name, args, original_command=None): # 関数の動的実行
        
        if func_name not in self.command_map:
            raise ValueError(f"[Excutor.execute] Unknown command: {func_name}")
        func = self.command_map[func_name] # 辞書command_mapのAutomatonのインスタンスへのポインタを使って関数を実行する
        result = func(args) # 関数名（ユーザー名称）を引数先頭に追加して実行
        
        # ログがあれば記録する
        if self.logger and original_command:
            self.logger.log(original_command, result)
        return result

    def execute_from_line(self, command_line):
        print(f"[Excutor.execute_from_line] comand_line : {command_line}")
        func_name, args = self.parser.parse(command_line)                             # パーサーで解読
        print(f"[Excutor.execute_from_line] fname = {func_name} , args = {args}")     
        return self.execute(func_name, args, original_command=command_line)

#c Scheduler           ★
class CommandScheduler:
    def __init__(self, operator):
        """
        operator: CommandOperator のインスタンス
        """
        self.operator = operator
        self.schedule = []  # [(datetime, command_str)] のリスト

    def add(self, scheduled_time, command_str):
        """スケジュールにコマンドを追加"""
        self.schedule.append((scheduled_time, command_str))

    def run(self):
        """
        登録されたコマンドを時刻順に監視しながら実行
        """
        print(f"[Scheduler.run] 監視開始: {len(self.schedule)} 件のコマンドをスケジュール")
        self.schedule.sort()  # 時刻順に並べる
        executed = set()  # 実行済みインデックス

        while len(executed) < len(self.schedule):
            now = datetime.now()
            for i, (scheduled_time, command_str) in enumerate(self.schedule):
                if i in executed:
                    continue
                if now >= scheduled_time:
                    print(f"[Scheduler.run] {now.strftime('%H:%M:%S')} 実行: {command_str}")
                    try:
                        result = self.operator.execute_from_line(command_str)
                        print(f"[Scheduler.run] > {command_str} => {result}")
                    except Exception as e:
                        print(f"[Scheduler.run] ! エラー: {e}")
                    executed.add(i)
            time.sleep(1)  # 1秒ごとに監視

#c pdfcraft                  ★
class PDFcraft : 
    def __init__(self, parser, operator, scheduler, automaton):
        self.parser = parser
        self.operator = operator
        self.scheduler = scheduler
        self.automaton = automaton

    def run(self, file_path):
        """
        commands.txt などを読み込み、即時／スケジュール分に分けて処理
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            
            for line in f:
                
                line = line.strip()
                
                if not line or line.startswith("#"):
                    print("continue")
        
                    continue  # 空行・コメント行は無視

                if line.startswith("at "):
                    # スケジュール付きコマンド
                    
                    try:
                        scheduled_time, command_str = self.parser.parse_scheduled(line)
                        self.scheduler.add(scheduled_time, command_str)
                        print(f"[pdfcraft.run] 予約登録: {scheduled_time.strftime('%H:%M:%S')} → {command_str}")
                    except Exception as e:
                        print(f"[pdfcraft.run] スケジュール解析エラー: {e}")
                else:
                    # 即時実行
                    try:
                        result = self.operator.execute_from_line(line)
                        print(f"[pdfcraft.run] 即時実行: {line} => {result}")
                    except Exception as e:
                        print(f"[pdfcraft.run] 即時実行エラー: {e}")

        # スケジュール実行を開始
        self.scheduler.run()
        
        return result


#c logger      ★ #PDFcraftﾚﾍﾞﾙのログ
class CommandLogger: 
    def __init__(self, log_path, isSaveText: bool = False):
        self.log_path = log_path
        self._flag_log_to_file = isSaveText  # ← ログ書き込みON/OFFフラグ

    def set_flag_logs_to_file(self, TrueOrFalse=True):  # ← 外部から変更可能に
        self._flag_log_to_file = TrueOrFalse

    def log(self, command_line, result=None):
        if not self._flag_log_to_file:
            return  # 書き込まないなら何もしない

        now = datetime.now().isoformat()
        with open(self.log_path, "a", encoding="utf-8") as f:
            if result is not None:
                f.write(f"{now} {command_line} => {result}\n")
            else:
                f.write(f"{now} {command_line}\n")

    def replay(self, operator):
        print("[Replay] 実行ログから再実行開始")
        with open(self.log_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    _, command_part = line.split(" ", 1)  # タイムスタンプを除去
                    command_only = command_part.split("=>")[0].strip()  # コマンド部分だけを抽出
                    result = operator.execute_from_line(command_only)
                    print(f"[Replay] {command_only} => {result}")
                except Exception as e:
                    print(f"[Replay Error] {line}: {e}")


#c FLogger             ★#PDFcraftの下のautomatonﾚﾍﾞﾙのログ
class FunctionLogger:  # 関数実行時のデバッグ用のログファイル出力
    def __init__(self, log_path, isSaveText: bool = False):  # ← スイッチ追加
        self.log_path = log_path
        self._logs = []
        self._flag_log_to_file = isSaveText  # ← フラグ初期化
        self._function_name = ""

    def set_flag_logs_to_file(self, TrueOrFalse: bool = True):
        self._flag_log_to_file = TrueOrFalse

    def set_function_name(self, fname):
        self._function_name = fname

    def log(self, msg):
        self.__add_log(msg)

    def read_log(self, index=-1):
        return self.__get_latest_log(index)

    def get_all_logs(self):
        return self._logs.copy()

    def __add_log(self, message: str):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        message = f"{timestamp}:[{self._function_name}] {message}"
        self._logs.append(message)
        if self._flag_log_to_file:
            self.__write_to_file(message)

    def __get_latest_log(self, no):
        return self._logs[no] if self._logs else None

    def __write_to_file(self, message: str):
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(message + '\n')


class QueueHandler:  # キュー操作を行うクラス★

    def shift(self, array, defaultvalue=""):  # Perlのshiftに近い関数。defaultvalueは省略可能
        if len(array) > 0:
            firstElement = array[0]
            del array[0]  # Pythonのリストは0-based index
            return firstElement
        else:
            return defaultvalue  # 空の場合はデフォルト値を返す

# === コマンド定義 ===
class Automaton: # 自動処理クラス★
    def __init__(self, queue_handler, logger):
        self.queue = queue_handler
        self.olog = logger

    # --- log 出力 ---
    def log(self, mystr):
        return self.olog.log(mystr)
        

    def convert_pdf_to_word(self, args): # PDFをWordに変換
        function_name = self.queue.shift(args)       # str 最初の引数は、自らの関数名
        input_pdf_path = self.queue.shift(args)      # str ワード化するPDFファイル
        output_docx_path = self.queue.shift(args)    # str ワードの出力先ファイル名（フルパス）

        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----       
        cv = Converter(input_pdf_path)                 # 変換器を作成
        cv.convert(output_docx_path, start=0, end=None)  # 全ページ変換
        cv.close()

        # Wordアプリケーションを開く
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = True  # Wordを表示する

        # 変換したWord文書を開く
        doc = word_app.Documents.Open(output_docx_path)

        # ここからWordでの操作を追加できます
        time.sleep(5)  # Wordが開くのを待つ
        try:
            # キー入力を送信する場合、pyautoguiを使う
            pyautogui.hotkey('alt', 'r')  # Alt + R
            time.sleep(0.3)
            pyautogui.hotkey('l')  # L
            time.sleep(0.3)
            pyautogui.hotkey('t')  # T
            time.sleep(3.0)
            # Tabを5回送信
            for _ in range(5):
                pyautogui.press('tab')
                time.sleep(0.3)  # Tabの間に待機時間を設ける

            # Enterキーを送信
            pyautogui.press('enter')

        except Exception as e:
            print(f"キー送信中にエラーが発生しました: {e}")

        # 必要に応じてドキュメントを保存して閉じる
        # doc.Save()  # 変更があれば保存
        # doc.Close()  # ドキュメントを閉じる
        # word_app.Quit()  # Wordアプリケーションを終了する

        #--- 結果 ----
        self.log(f"ワード化完了 {input_pdf_path} をワード化し、 {output_docx_path} に保存しました。")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    def copy_ppt_watermark(self, args):
        function_name = self.queue.shift(args)       # str 最初の引数は、自らの関数名
        input_ppt_path = self.queue.shift(args)      # str コピーするPPTファイル
        last_page = self.queue.shift(args)           # int
        copy_elements = self.queue.shift(args)       # []  PPT１ページ目コピーする透かしの要素名
        overwrite_elements = self.queue.shift(args)  # []  PPT１ページ目上書きするテキストの要素名
        output_pdf_path = self.queue.shift(args)     # str PDFの出力先ファイル名（フルパス）

        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        if not input_ppt_path:
            return

        # まず開いているプレゼン一覧から一致するものを探す
        pythoncom.CoInitialize()
        pptApp = win32com.client.Dispatch("PowerPoint.Application")
        target_instance = None

        for pres in pptApp.Presentations:
            # フルパスが一致すればそれを使う
            if pres.FullName.lower() == input_ppt_path.lower():
                target_instance = pres
                break

        # 開いてなければ新たに開く
        if not target_instance:
            try:
                target_instance = pptApp.Presentations.Open(input_ppt_path, WithWindow=False)
            except Exception as e:
                print(f"PowerPoint ファイルを開けませんでした: {e}")
                return

        # コピーして作業
        copy_instance = self._ppt_copy_and_open_powerpoint(target_instance.FullName)
        if copy_instance:
            self._ppt_apply_watermark_and_export_pdf(copy_instance, last_page, copy_elements, overwrite_elements)            

        #--- 結果 ----
        self.log(f"透かしコピー完了 {input_ppt_path} の透かしをコピーし、 {output_pdf_path} に保存しました。")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }


    def _ppt_retrieve_or_open_powerpoint(self, ppt_file_path):
        #指定されたPowerPointファイルのインスタンスを取得する関数"""
        pythoncom.CoInitialize()

        try:
            pptApp = win32com.client.Dispatch("PowerPoint.Application")

            # プレゼンテーションがすでに開いているか確認
            for presentation in pptApp.Presentations:
                if presentation.FullName == ppt_file_path:
                    return presentation  # 開いているインスタンスを返す

            # 開いていない場合は新たに開く
            return pptApp.Presentations.Open(ppt_file_path, WithWindow=False)

        except Exception as e:
            messagebox.showerror("エラー", f"指定されたPowerPointファイルを開けませんでした。\n\n{e}")
            return None

    def _ppt_copy_and_open_powerpoint(self, ppt_file_path):
        #指定されたPowerPointファイルをコピーする関数"""
        original = self._ppt_retrieve_or_open_powerpoint(ppt_file_path)
        if not original:
            return None

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        original_name = os.path.splitext(original.Name)[0]
        copy_path = os.path.join(desktop, f"{original_name}_watermark.pptx")

        original.SaveCopyAs(copy_path)  # コピーを保存
        return original.Application.Presentations.Open(copy_path, WithWindow=False)  # コピーしたプレゼンテーションを返す


    def _ppt_apply_watermark_and_export_pdf(self, ppt_instance, last_page, copy_element_name, overwrite_element_name):
        #コピーしたPowerPointインスタンスに対してcopy_element_nameとoverwrite_element_nameを全ページにコピーし、last_pageまでを保存する関数"""
        if not last_page:
            ppt_instance.Close()
            return 0, 0

        source_slide = ppt_instance.Slides(1)

        for shape in source_slide.Shapes:
            if shape.Name in overwrite_element_name:
                shape.TextFrame.TextRange.Text = f"/ {last_page}" # 総ページ要素
            if shape.Name in copy_element_name:
                shape.Copy() # 透かし要素
                for i in range(2, ppt_instance.Slides.Count + 1):
                    ppt_instance.Slides(i).Shapes.Paste()

        for i in range(ppt_instance.Slides.Count, last_page, -1):
            ppt_instance.Slides(i).Delete()

        pdf_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{os.path.splitext(ppt_instance.Name)[0]}.pdf")
        ppt_instance.SaveAs(pdf_path, 32)
        ppt_instance.Save() #コピーしたパワポも保存
        ppt_instance.Close()

        return last_page, pdf_path


    # --- 透かしpdfを作る関数 ---
    def _get_pdf_page_size(self, pdf_path):
        reader = PdfReader(pdf_path)
        page = reader.pages[0]
        media_box = page.mediabox
        width = float(media_box.upper_right[0]) - float(media_box.lower_left[0])
        height = float(media_box.upper_right[1]) - float(media_box.lower_left[1])
        return width, height

    def watermark_pdf(self, args):
        function_name = self.queue.shift(args)       # str 関数名
        pdf_paths = self.queue.shift(args)           # list 透かしを入れるPDFのリスト
        text = self.queue.shift(args)                # str 透かし文字
        output_paths = self.queue.shift(args)        # list 出力先ファイル名のリスト

        if isinstance(pdf_paths, str):
            pdf_paths = [pdf_paths]

        if isinstance(output_paths, str):
            output_paths = [output_paths]
            
        self.olog.set_function_name(function_name)
        self.log("透かし処理を開始します")

        for pdf_path, output_path in zip(pdf_paths, output_paths):
            self.log(f"{pdf_path} に透かしを追加 → {output_path}")

            width, height = self._get_pdf_page_size(pdf_path)
            font_size = 40
            watermark_pdf_path = os.path.splitext(output_path)[0] + "_temp_watermark.pdf"

            # 透かしページを作成（reportlab）
            c = canvas.Canvas(watermark_pdf_path, pagesize=(width, height))
            c.setFont("Helvetica", font_size)
            c.setFillGray(0.5, 0.5)
            x_gap = int(font_size * 8.0)
            y_gap = int(font_size * 8.0)
            w = len(text)

            for y in range(0, int(height + y_gap), y_gap):
                for x in range(0, int(width + x_gap), int(w*font_size/4) + x_gap):
                    c.saveState()
                    c.translate(x, y)
                    c.rotate(45)
                    c.drawCentredString(0, 0, text)
                    c.restoreState()

            c.save()

            # PDFに透かしを適用
            reader = PdfReader(pdf_path)
            watermark_reader = PdfReader(watermark_pdf_path)
            writer = PdfWriter()

            watermark_page = watermark_reader.pages[0]

            for page in reader.pages:
                page.merge_page(watermark_page)
                writer.add_page(page)

            with open(output_path, "wb") as f:
                writer.write(f)

            os.remove(watermark_pdf_path)  # 一時ファイルを削除

        self.olog.set_function_name("")
        return {
            "function_name": function_name,
            "status": 1,
            "message": f"{len(pdf_paths)}個のPDFに透かしを追加しました。",
            "log": self.olog.get_all_logs()
        }


    # --- PDF 結合処理 ---
    def merge_pdfs(self, args):  # 複数のPDFを結合

        function_name = self.queue.shift(args)       # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        pdf_paths = self.queue.shift(args)           # [] 結合するPDFのファイル名リスト（フルパス）
        output_path = self.queue.shift(args)         # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDF結合を開始します ---")  
        merger = PdfMerger()                         # PyPDFのマージャーを生成
        for pdf in pdf_paths:                        # 各PDFを順に処理
            filename = os.path.splitext(os.path.basename(pdf))[0]  # ファイル名のみ取得
            merger.append(pdf, outline_item=filename)             # PDFにしおりを付けて追加

        merger.write(output_path)                    # マージしたPDFを書き込み
        merger.close()                               # マージャーを閉じる

        #--- 結果 ----
        self.log(f"ファイルを結合しました:\n{output_path}")           # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    
    def add_jpg_to_pdf(self, args): #1つのPDFに複数の画像ファイルを追加
        
        function_name = self.queue.shift(args)       # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_pdf_path = self.queue.shift(args)      # str 結合するベースのPDFのファイル名（フルパス）
        input_jpg_paths = self.queue.shift(args)     # [] 結合する画像ファイル名のリスト（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定
        
        #--- 処理 ----
        self.log(f"起動フォルダに複数の画像ファイルがあります。画像をPDFファイルに追加します")
        self.log("add_jpg_to_pdf()を開始します。（画像ファイルのPDF変換）")
        print ("-----------add_date_to_filename-------------")
        output_pdf_path2 =  add_date_to_filename(input_pdf_path) #末尾に日付付加
        print (f"出力先：{output_pdf_path2}")
        self.jpg_to_pdf(["jpg_to_pdf", input_jpg_paths, output_pdf_path2])
  
        self.log("merge_pdfs()を開始します。（元のPDFファイルと、作成した画像PDFファイルの結合）")
        input_pdf_paths = [input_pdf_path]
        input_pdf_paths.append(output_pdf_path2) 
        output_pdf_path1 =  os.path.dirname(input_jpg_paths[0]) + "\\merge_jpgspdf.pdf" 
        self.merge_pdfs(["merge_pdfs", input_pdf_paths, output_pdf_path1])
        
        #名前の変更、中間ファイル削除
        self.log("元のPDFファイル名に日付をつけてバックアップ保存します")
        shutil.move(input_pdf_path, add_date_to_filename(input_pdf_path)) 
        self.log("生成したPDFファイルの名前を、元ファイルの名前に変更します")
        shutil.move(output_pdf_path1, input_pdf_path)
        if os.path.exists(output_path1):
            self.log("中間ファイルを削除します")
            os.remove(output_pdf_path2)
        
        #--- 結果 ----
        self.log(f"完了！ 画像ファイルをpdfに追加ました：\n{input_pdf_path}")            # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    def jpg_to_pdf(self, args): #画像ファイルをpdfにする
        
        function_name = self.queue.shift(args)       # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_image_paths = self.queue.shift(args)     # [] 結合する画像ファイル名のリスト（フルパス）
        output_pdf_path = self.queue.shift(args)     # str 出力するPDFファイル名（フルパス）
    
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定

        #--- 処理 ----
        image_list = []
        for path in input_image_paths:
            img = Image.open(path).convert("RGB")
            image_list.append(img)

        if image_list:
            first = image_list[0]
            others = image_list[1:]
            first.save(output_pdf_path, save_all=True, append_images=others)

        #--- 結果 ----
        self.log(f"画像を {output_pdf_path} に保存しました。")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    # --- PDF 分割処理 ---
    def split_pdf(self, args): #１つのPDFを２つに分割

        function_name = self.queue.shift(args)        # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_path = self.queue.shift(args)           # str 分割するファイル名（フルパス）
        split_at = int(self.queue.shift(args))        # str 分割するページ数
        output_path1 = self.queue.shift(args)         # str 出力先ファイル名（フルパス）
        output_path2 = self.queue.shift(args)         # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDF分割を開始します ---")
        reader = PdfReader(input_path)
        total_pages = len(reader.pages)

        if split_at < 1 or split_at >= total_pages:
            raise ValueError(f"1?{total_pages - 1} の範囲で入力してください。")

        writer1 = PdfWriter()
        writer2 = PdfWriter()

        for i in range(split_at):
            writer1.add_page(reader.pages[i])
        for i in range(split_at, total_pages):
            writer2.add_page(reader.pages[i])

        with open(output_path1, "wb") as f1:
            writer1.write(f1)
        with open(output_path2, "wb") as f2:
            writer2.write(f2)
        
        #--- 結果 ----
        self.log(f"1～{split_at}、{split_at+1}～{total_pages}ページに分割しました")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }
    
    # --- PDF 置換処理 ---
    def replace_pdf(self, args):

        function_name = self.queue.shift(args)        # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_path1 = self.queue.shift(args)          # str 置換されるファイル名（フルパス）
        split_start = int(self.queue.shift(args))     # str 置換されるページ数　始まり
        split_end = int(self.queue.shift(args))       # str 置換されるページ数　終わり
        input_path2 = self.queue.shift(args)          # str 置換するファイル名（フルパス）
        output_path = self.queue.shift(args)          # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDF置換を開始します ---")
        reader = PdfReader(input_path1)
        total_pages = len(reader.pages)

        if split_start < 1 or split_start >= total_pages:
            raise ValueError(f"開始ページは、1～{total_pages} の範囲で入力してください。")

        if split_end < split_end or split_end > total_pages:
            raise ValueError(f"終了ページは、{split_start}～{total_pages} の範囲で入力してください。")
        
        folder_name = os.path.basename(input_path1)
        if split_start == 1:
            self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", folder_name + "C.pdf"])
            #self.split_pdf(["split_pdf", folder_name + "AB.pdf", split_start-1, folder_name + "A.pdf", folder_name + "B.pdf"])
            self.merge_pdfs(["merge_pdfs", [input_path2, folder_name + "C.pdf"], output_path])
            os.remove(folder_name + "AB.pdf") # 中間ファイル削除
            os.remove(folder_name + "C.pdf") # 中間ファイル削除
        
        elif split_end == total_pages:
            #self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", folder_name + "C.pdf"])
            self.split_pdf(["split_pdf", input_path1, split_start-1, folder_name + "A.pdf", folder_name + "B.pdf"])
            self.merge_pdfs(["merge_pdfs", [folder_name + "A.pdf", input_path2], output_path])
            os.remove(folder_name + "A.pdf") # 中間ファイル削除
            os.remove(folder_name + "B.pdf") # 中間ファイル削除
        
        else:
            self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", folder_name + "C.pdf"])
            self.split_pdf(["split_pdf", folder_name + "AB.pdf",  split_start-1, folder_name + "A.pdf", folder_name + "B.pdf"])
            self.merge_pdfs(["merge_pdfs", [folder_name + "A.pdf", input_path2, folder_name + "C.pdf"], output_path])
            os.remove(folder_name + "AB.pdf") # 中間ファイル削除
            os.remove(folder_name + "A.pdf") # 中間ファイル削除
            os.remove(folder_name + "B.pdf") # 中間ファイル削除
            os.remove(folder_name + "C.pdf") # 中間ファイル削除
        
        #--- 結果 ----
        self.log(f"{split_start}～{split_end}ページを置換しました")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    # --- PDF 挿入処理 ---
    def insert_pdf(self, args):

        function_name = self.queue.shift(args)        # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_path1 = self.queue.shift(args)          # str 挿入されるファイル名（フルパス）
        insert_page_no = int(self.queue.shift(args))  # str 挿入ページ　このページの手前に挿入される
        input_path2 = self.queue.shift(args)          # str 挿入するファイル名（フルパス）
        output_path = self.queue.shift(args)          # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDF挿入を開始します ---")
        reader = PdfReader(input_path1)
        total_pages = len(reader.pages)

        if insert_page_no < 1:
            raise ValueError(f"開始ページは、1～{total_pages} の範囲で入力してください。")

        folder_name = os.path.basename(input_path1)
        if insert_page_no == 1:
            self.merge_pdfs(["merge_pdfs", [input_path2, input_path1], output_path])
        
        elif insert_page_no > total_pages:
            self.merge_pdfs(["merge_pdfs", [input_path1, input_path2], output_path])
        
        else:
            self.split_pdf(["split_pdf", input_path1, insert_page_no-1, folder_name + "A.pdf", folder_name + "B.pdf"])
            self.merge_pdfs(["merge_pdfs", [folder_name + "A.pdf", input_path2, folder_name + "B.pdf"], output_path])
            os.remove(folder_name + "A.pdf") # 中間ファイル削除
            os.remove(folder_name + "B.pdf") # 中間ファイル削除
        
        #--- 結果 ----
        self.log(f"{input_path2}の{insert_page_no}ページの前に、 {input_path2}を挿入しました")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }
    
    # --- PDF ページ削除処理 ---
    def remove_pdf(self, args):

        function_name = self.queue.shift(args)        # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_path1 = self.queue.shift(args)          # str 置換されるファイル名（フルパス）
        split_start = int(self.queue.shift(args))     # str 置換されるページ数　始まり
        split_end = int(self.queue.shift(args))       # str 置換されるページ数　終わり
        output_path = self.queue.shift(args)          # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDFページ削除を開始します ---")
        reader = PdfReader(input_path1)
        total_pages = len(reader.pages)

        if split_start < 1 or split_start >= total_pages:
            raise ValueError(f"開始ページは、1?{total_pages} の範囲で入力してください。")

        if split_end < split_end or split_end > total_pages:
            raise ValueError(f"終了ページは、{split_start}?{total_pages} の範囲で入力してください。")
        
        folder_name = os.path.basename(input_path1)
        if split_start == 1:
            self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", output_path])
            os.remove(folder_name + "AB.pdf") # 中間ファイル削除
        
        elif split_end == total_pages:
            self.split_pdf(["split_pdf", input_path1, split_start-1, output_path, folder_name + "B.pdf"])
            os.remove(folder_name + "B.pdf") # 中間ファイル削除

        else:
            self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", folder_name + "C.pdf"])
            self.split_pdf(["split_pdf", folder_name + "AB.pdf",  split_start-1, folder_name + "A.pdf", folder_name + "B.pdf"])
            self.merge_pdfs(["merge_pdf", [folder_name + "A.pdf", folder_name + "C.pdf"], output_path])
            os.remove(folder_name + "AB.pdf") # 中間ファイル削除
            os.remove(folder_name + "A.pdf") # 中間ファイル削除
            os.remove(folder_name + "B.pdf") # 中間ファイル削除
            os.remove(folder_name + "C.pdf") # 中間ファイル削除
            
        #--- 結果 ----
        self.log(f"{split_start}～{split_end}ページを削除しました")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }
        
    # --- PDF 抽出処理 ---
    def extract_pdf(self, args):

        function_name = self.queue.shift(args)        # str 最初の引数は、自らの関数名。関数実行中のlog出力で[関数名]を出力するため
        input_path1 = self.queue.shift(args)          # str 抽出されるファイル名（フルパス）
        split_start = int(self.queue.shift(args))     # str 抽出されるページ数　始まり
        split_end = int(self.queue.shift(args))       # str 抽出されるページ数　終わり
        output_path = self.queue.shift(args)          # str 出力先ファイル名（フルパス）
        
        self.olog.set_function_name(function_name)    # FunctionLoggerクラスに関数名を設定                      
 
        #--- 処理 ----
        self.log(f"--- PDF抽出処理を開始します ---")
        reader = PdfReader(input_path1)
        total_pages = len(reader.pages)

        folder_name = os.path.basename(input_path1)
        if split_start < 1 or split_start > total_pages or split_end - split_start + 1 == total_pages:
            self.log(f"{os.path.basename(input_path1)} の処理中に、無効な値が入力されました")
            return {
                "function_name": function_name,
                "status": -1,
                "message": self.log(-1),
                "log": self.olog.get_all_logs()
            }
        
        elif split_start == 1:
            self.split_pdf(["split_pdf", input_path1, split_end, output_path,  folder_name + "B.pdf"])
            os.remove(folder_name + "B.pdf") # 中間ファイル削除
        elif split_end == total_pages:
            self.split_pdf(["split_pdf", input_path1, split_start-1, folder_name + "A.pdf", output_path])
            os.remove(folder_name + "A.pdf") # 中間ファイル削除                    
        else:
            self.split_pdf(["split_pdf", input_path1, split_end, folder_name + "AB.pdf", folder_name + "C.pdf"])
            self.split_pdf(["split_pdf", folder_name + "AB.pdf",  split_start-1, folder_name + "A.pdf", output_path])
            os.remove(folder_name + "AB.pdf") # 中間ファイル削除
            os.remove(folder_name + "A.pdf") # 中間ファイル削除
            os.remove(folder_name + "C.pdf") # 中間ファイル削除
               
        
        #--- 結果 ----
        self.log(f"{split_start}～{split_end}ページを抽出しました")         # 成功ログを追加
        self.olog.set_function_name("")             # FunctionLoggerクラスに関数名を削除 
        return {
            "function_name": function_name, 
            "status": 1, 
            "message": self.olog.read_log(-1), 
            "log": self.olog.get_all_logs()
        }

    def password_pdfs(self, args):
        function_name = self.queue.shift(args)           # str 関数名
        pdf_paths = self.queue.shift(args)               # list PDFファイル群
        password = self.queue.shift(args)                # str パスワード
        output_paths = self.queue.shift(args)            # list 出力ファイル名リスト

        self.olog.set_function_name(function_name)

        if isinstance(pdf_paths, str):
            pdf_paths = [pdf_paths]

        if isinstance(output_paths, str):
            output_paths = [output_paths]

        if not pdf_paths or not output_paths or len(pdf_paths) != len(output_paths):
            msg = "入力PDFと出力パスの数が一致しません"
            self.log(msg)
            return {
                "function_name": function_name,
                "status": -1,
                "message": msg,
                "log": self.olog.get_all_logs()
            }

        new_files = []
        new_filenames = []

        for path, out_path in zip(pdf_paths, output_paths):
            try:
                reader = PdfReader(path)
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)

                writer.encrypt(password)

                with open(out_path, "wb") as f:
                    writer.write(f)

                self.log(f"{os.path.basename(path)} にパスワードを付与 → {out_path}")
                new_files.append(out_path)
                new_filenames.append(os.path.basename(out_path))

            except Exception as e:
                self.log(f"エラー発生：{e}")
                return {
                    "function_name": function_name,
                    "status": -1,
                    "message": f"{os.path.basename(path)} の処理中にエラーが発生しました：\n\n{e}",
                    "log": self.olog.get_all_logs()
                }

        self.log("全ファイルにパスワードを付与しました。")
        self.olog.set_function_name("")

        return {
            "function_name": function_name,
            "status": 1,
            "message": f"{len(new_files)}個のファイルにパスワードを付与しました。",
            "log": self.olog.get_all_logs(),
            "output_files": new_files,
            "output_filenames": new_filenames,
            "output_dir": os.path.dirname(new_files[0]) if new_files else ""
        }


# --- 文字列変換関数 ---
def add_date_to_filename(path):
    """
    指定されたファイルパスに、当日の日付（yymmdd）を付加して新しいパスを返す。
    例: 'file.pdf' → 'file_240322.pdf'
    """
    folder = os.path.dirname(path)  # フォルダ部分を取得
    base, ext = os.path.splitext(os.path.basename(path))  # ファイル名と拡張子に分割
    today_str = datetime.now().strftime("%y%m%d")  # 今日の日付を yymmdd 形式で取得
    new_name = f"{base}_{today_str}{ext}"  # 日付付きの新しいファイル名を作成
    return os.path.join(folder, new_name)  # フルパスとして返す

def read_commandstxt(filepath):
    lines = []  # 結果をためるリスト
    with open(filepath, "r", encoding="utf-8") as f:  # ファイルをUTF-8で開く
        for line in f:  # 1行ずつ読む
            line = line.strip()  # 前後の空白や改行を削除
            if not line or line.startswith("#"):  # 空行またはコメント行ならスキップ
                continue
            lines.append(line)  # 有効な行だけ追加
    return "\n".join(lines)  # すべての行を \n 区切りで連結して返す

def read_commands_with_substitution(filepaths):
    
    #コマンドテキストを読み込み、以下の機能に対応：
    """
    1. 定数定義： x1 = "C:/example.pdf"
    2. ユーザー入力： x2 = ?
    3. 説明付き入力：説明: x3 = ?
    4. 説明文に「パス」「path」「アドレス」が含まれていれば、ファイル選択ダイアログで入力
    5. 複数ファイル対応（filepathsはlist）
    6. 入力された値が数値ならクオートなし、文字列ならクオート付き
    """
    #想定するテキストファイル 
    """
    結合ファイルNo1: x1 = ?
    PDF化画像No2 : x2 = ?
    # 結合
    merge([x1, "C:/Users/niwakenji/Desktop/python/input2.pdf", "C:/Users/niwakenji/Desktop/python/input1.pdf"], "C:/Users/niwakenji/Desktop/python\merge.pdf")
    # 画像PDF化
    pdf(["C:/Users/niwakenji/Desktop/python/新しいフォルダー (3)/241001_免許_表.jpg", x2, "C:/Users/niwakenji/Desktop/python/新しいフォルダー (3)/250215_ラクマ_マウス購入したものと異なる.jpg"], "C:/Users/niwakenji/Desktop/python/新しいフォルダー (3)\image_250327.pdf")
    """    

    all_lines = []
    root = Tk()
    root.withdraw()

    for filepath in filepaths:
        substitutions = {}

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            with open(filepath, "r", encoding="cp932") as f:
                lines = f.readlines()

        for line in lines:
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            # --- 変数定義（= を含み、かつ if ではない） ---
            if "=" in line and not line.startswith("if"):
                try:
                    # 説明つき（「:」区切り）の場合を分離
                    if ":" in line:
                        label_part, expr_part = line.split(":", 1)
                        expr_part = expr_part.strip()
                        description = label_part.strip()
                    else:
                        expr_part = line
                        description = None

                    var, value = expr_part.split("=", 1)
                    var = var.strip()
                    value = value.strip().strip('"')

                    # --- ユーザー入力モード ---
                    if value == "?":
                        prompt = f"{var} の値を入力してください："
                        if description:
                            prompt = f"{description}: {var} の値を入力してください："

                        # 説明文にパス系の語句が含まれていればファイル選択ダイアログを使用
                        if description and re.search(r"パス|path|アドレス|ファイル|在りか", description, re.IGNORECASE):
                            user_value = filedialog.askopenfilename(title=prompt)
                        else:
                            user_value = simpledialog.askstring("変数の入力", prompt)

                        if user_value is None:
                            user_value = ""

                        # 数値ならそのまま、文字列ならクオート付きにする
                        if re.fullmatch(r'[0-9.]+', user_value):  # 数値っぽい
                            value = user_value
                        else:
                            value = f'"{user_value}"'

                    substitutions[var] = value
                except Exception as e:
                    print(f"[ERROR] 変数定義パース失敗: {line} ({e})")
                    pass

            else:
                # --- 通常コマンド：変数をすべて置換 ---
                for key, val in substitutions.items():
                    line = line.replace(f'"{key}"', val)  # "x1" のような文字列を先に
                    line = re.sub(rf'\b{re.escape(key)}\b', val, line)  # x1 のような変数単体も置換
                all_lines.append(line)

    root.destroy()
    return "\n".join(all_lines)



# main関数 === ★
if __name__ == "__main__":
    
    # 起動画面の終了
    _loading_win.destroy()  # ← 終わったら閉じる
    
    # PDFcraftクラス(parser + operator + scheduler + automaton) の構築
    automatonlog_filename = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "log_automaton.txt") # アドレスの区切りを自動処理
    automatonlogger = FunctionLogger(automatonlog_filename,False)
    queuehandler = QueueHandler()
    automaton = Automaton(queuehandler, automatonlogger)   

    parser = CommandParser()
    pdfcraftlog_filename = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "log_pdfcraft.txt")
    logger = CommandLogger(pdfcraftlog_filename,False)
    operator = CommandOperator(parser, logger)
    scheduler = CommandScheduler(operator)
    pdfcraft = PDFcraft (parser, operator, scheduler, automaton)

    # コマンドマップの設定
    command_map = { # 辞書に、各コマンド名に対応するメソッドへの参照を格納する commands.txtで、merge(...)と書くと、右の関数が実行される
        "merge": pdfcraft.automaton.merge_pdfs,  
        "split": pdfcraft.automaton.split_pdf,  
        "replace": pdfcraft.automaton.replace_pdf,  
        "remove": pdfcraft.automaton.remove_pdf,  
        "extract": pdfcraft.automaton.extract_pdf,
        "convert": pdfcraft.automaton.convert_pdf_to_word,
        "watermark": pdfcraft.automaton.watermark_pdf,
        "sukashi": pdfcraft.automaton.copy_ppt_watermark,
        "add": pdfcraft.automaton.add_jpg_to_pdf,
        "pdf": pdfcraft.automaton.jpg_to_pdf,
        "insert": pdfcraft.automaton.insert_pdf,
        "password": pdfcraft.automaton.password_pdfs,
    }
    pdfcraft.operator.set_command_map(command_map) # コマンドマップをオペレータに伝える

    # CommandMakerクラスの構築と起動
    commandmaker = CommandMaker(pdfcraft)
    commandmaker.run()


