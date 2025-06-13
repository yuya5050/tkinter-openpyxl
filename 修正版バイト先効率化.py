#ファイルを開く→名前エントリー画面を表示→
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from datetime import datetime, timedelta
import xlwings as xw # xlwingsライブラリのインポート

#enntry_window→submit_name→open_window→select_file→submit_file→file_parse（excel_serial_to_date)→loopaction→next_action→excel_input



def clear_widgets():
    for widget in root.winfo_children():
        widget.destroy()

def entry_window():
    global name, entry, submit_button, root, file_path, select_button
    root = tk.Tk()
    root.title("名前の提出")
    root.geometry("300x500")

    # 名前入力欄
    entry = tk.Entry(root)
    entry.pack(pady=10)

    # 提出ボタン
    submit_button = tk.Button(root, text="提出", command=submit_name)
    submit_button.pack()






def submit_name():
    global name
    name = ""
    # 入力された名前を取得してグローバル変数に代入
    name = entry.get()
    print(f"提出された名前: {name}")
    # 入力欄とボタンを削除
    entry.destroy()
    submit_button.destroy()
    open_window()
    
def select_action():
    root.title(f"動作を選択({name})")
    button1=tk.Button(root, text="hello world", command=tanomo_action)
    button1.pack(pady=10)
    button2=tk.Button(root, text="サンプル使用", command=sample_action)
    button2.pack(pady=10)
    button3=tk.Button(root, text="その他　詳細をG列に記載", command=other_action)
    button3.pack(pady=10)
    button4=tk.Button(root, text="店舗間移動", command=move_action)
    button4.pack(pady=10)
    button5=tk.Button(root, text="破損", command=break_action)
    button5.pack(pady=10)
    label= tk.Label(root, text=f"選択されたファイル: {file_path}")
    label.pack(pady=10)
global last_list
last_list = []


def tanomo_action():
    global last_list, tanomo_list
    last_list = tanomo_list
    clear_widgets()
    loopaction()


def sample_action():
    global last_list, sample_list
    last_list = sample_list
    clear_widgets()
    loopaction()

def other_action():
    global last_list, other_list
    last_list = other_list
    clear_widgets()
    loopaction()
    
def move_action():
    global last_list, move_list
    last_list = move_list
    clear_widgets()
    loopaction()

def break_action():
    global last_list, break_list
    last_list = break_list
    clear_widgets()
    loopaction()
# ウィンドウの作成




def select_file():
    
    global file_path
    # Excelファイルだけを対象にファイル選択ダイアログを表示
    file_path = ""
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Excelファイルを選択"
    )
    if file_path:
        print(f"選択されたファイル: {file_path}")

def submit_file():
    if file_path:
        print(f"提出されたファイルパス: {file_path}")
        select_button.destroy()
        submit_button.destroy()
        file_parse()

def open_window():
    clear_widgets()
    # ファイル選択ボタン
    global select_button, submit_button
    root.title(f"ファイルを選択({name})")
    select_button = tk.Button(root, text="ファイルを選択", command=select_file)
    select_button.pack(pady=10)

    # 提出ボタン
    submit_button = tk.Button(root, text="提出", command=submit_file)
    submit_button.pack(pady=10)




def file_parse():
    try:
        global app,wb,ws
        app = xw.App(visible=False) # Excelアプリケーションを起動（Trueで表示、Falseで非表示）
        wb = app.books.open(file_path) # 指定されたパスのワークブックを開く
        ws = wb.sheets.active # アクティブなシートを取得

        rows = 0
        row_num = 6
        while True:
            cell_value = ws.range(f"B{row_num}").value # B列のセルの値を取得
            if cell_value is None:
                break # セルが空になったらループを終了
            rows += 1
            row_num += 1

        data_list = []
        row_number = 6
        for _ in range(rows):
            columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
            data_dict = {}
            for col in columns:
                cell_value = ws.range(f"{col}{row_number}").value # 各列のセルの値を取得
                data_dict[col] = cell_value
            data_dict["row_number"] = row_number # 行番号を辞書に追加
            data_list.append(data_dict)
            row_number += 1
        
        wb.close() # ワークブックを閉じる
        app.quit() # Excelアプリケーションを終了

        global all_list
        all_list = []
        for data in data_list:
            if data["K"] is None and data["J"] is None: # K列とJ列が空の行のみを対象
                all_list.append(data)
                
        global tanomo_list, sample_list, other_list, break_list, move_list
        tanomo_list = []
        sample_list = []    
        other_list = []
        break_list = []
        move_list = []
        
        for data in all_list:
            if data["H"] == 'hello world':
                tanomo_list.append(data)
            elif data["H"] == 'サンプル使用':
                sample_list.append(data)
            elif data['H'] == 'その他　詳細をG列に記載':
                other_list.append(data)
            elif data['H'] == '店舗間移動\u3000\u3000\u3000届け先をG列に記載→':
                move_list.append(data)
            elif data['H'] == '破損':
                break_list.append(data)
        select_action()

    except Exception as e:
        clear_widgets()
        label=tk.Label(root,text="エラー：他プログラムで同じエクセルファイルをひらいていませんか？")
        button=tk.Button(root,text="ファイル選択画面へ",command=open_window)
        label.pack(pady=10)
        button.pack(pady=10)
        

    
    
    
   
counter=0

def excel_serial_to_date(serial_number):
    base_date = datetime(1899, 12, 30)  # Excelは1900年1月1日を「1」と数える
    return base_date + timedelta(days=serial_number)

def loopaction():
    clear_widgets()
    global counter, input_number, date, sender, receiver, how_many,button,root, button2
    if len(last_list) == 0:
        label = tk.Label(root, text="処理するデータがありません。")
        label.pack(pady=10)
        root.title(f"処理完了({name})")
        next_button = tk.Button(root, text="他のアクションを選択", command=end_action)
        next_button.pack(pady=10)
    else:
        real_date=last_list[counter]["D"]
        date = tk.Label(root, text=f"移動日: {real_date.strftime("%Y-%m-%d")}")
        date.pack()
        sender = tk.Label(root, text=f"店舗: {last_list[counter]["B"]}")
        sender.pack()
        item_name = tk.Label(root, text=f"商品名: {last_list[counter]["F"]}")
        item_name.pack()
        item_number= tk.Label(root, text=f"商品番号: {last_list[counter]["E"]}")
        item_number.pack()
        receiver = tk.Label(root, text=f"使用用途: {last_list[counter]["I"]}")
        receiver.pack()
        how_many = tk.Label(root, text=f"送った個数: {last_list[counter]["G"]}")
        how_many.pack()
        global input_number, button
        input_number = tk.Entry(root)
        input_number.pack(pady=10)
        button = tk.Button(root, text="次へ", command=next_action)
        button.pack(pady=10)
        button2=tk.Button(root,text="飛ばす",command=skip_acton)
        button2.pack(pady=10)
        root.title(f"データ入力（処理番号)({name})")

def next_action():
    global counter, date, sender, receiver, how_many, input_number,button, last_list, button2,input_value
    input_value = input_number.get()
        #ここにエクセル入力する関数を追加する
    if input_value=="":
        loopaction()
    else:
        excel_input()

        clear_widgets()
        counter += 1
        if counter < len(last_list):
            
            loopaction()
        
        else:
            label= tk.Label(root, text="全てのデータを処理しました。")
            label.pack(pady=10)
            root.title("処理完了")
            next_button = tk.Button(root, text="次のアクションを選択", command=end_action)
            next_button.pack(pady=10)

def skip_acton():
    global counter, date, sender, receiver, how_many, input_number,button, last_list,button2

    #ここにエクセル入力する関数を追加する
    

    clear_widgets()
    counter += 1
    if counter < len(last_list):
        
        loopaction()
    
    else:
        label= tk.Label(root, text="全てのデータを処理しました。")
        label.pack(pady=10)
        root.title("処理完了")
        next_button = tk.Button(root, text="次のアクションを選択", command=end_action)
        next_button.pack(pady=10)
def end_action():
    global counter, last_list
    counter = 0
    clear_widgets()
    file_parse()
    
    



def excel_input():
    global input_number, last_list, counter, name
    
    app = None 
    
    try:
        
        
        ws = wb.sheets.active

       
        
        row_number = last_list[counter]["row_number"]
        
        ws.range(f"K{row_number}").value = input_value
        ws.range(f"J{row_number}").value = name
        
        wb.save()
        
        print(f"行 {row_number} の 処理番号に '{input_value}' を入力しました。")
        
    except Exception as e:
        print("エラーが発生しました:", e)
    finally:
        if app is not None and app.alive:
            wb.close()
            app.quit()


        
    






# メインループ開始

if __name__ == "__main__":
    entry_window()
    root.mainloop()


