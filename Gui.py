import tkinter as tk
from tkinter import ttk
import pandas as pd

import pptx_auto_gen
import xlsx_auto_gen

class MainWindow(tk.Frame):
    def __init__(self, master=None, parent=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("350x700")
        self.master.title("メニュー名入力したら色々できてるやつ")
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        self.set_data()
        self.create_widgets()
    
    def create_widgets(self):
        self.pw_main = ttk.PanedWindow(self.master, orient="vertical")
        self.pw_main.pack(expand=True, fill=tk.BOTH, side="left")
        self.pw_top = ttk.PanedWindow(self.pw_main, orient="horizontal")
        self.pw_main.add(self.pw_top)
        self.tmp(self.pw_top)

    def set_data(self):
        self.df_sales_management = pd.read_excel("../output/sales_management.xlsx", sheet_name="販売個数", header=0, index_col=None)
        self.colname_list = ["日付", "弁当名等", "搬入個数"]  # 結果に表示させる列名
        self.width_list = [100, 200, 50]
        self.search_col = "日付"  # 検索キーワードの入力されている列名

        self.df_menu_master = pd.read_excel("../output/sales_management.xlsx", sheet_name="弁当名マスタ", header=0, index_col=None)

    def btn_click(self):
        menu_array = []
        quantity_array = []
        if self.txt1.get() != "" :
            menu_array.append(self.txt1.get())
            quantity_array.append(int(self.quantity1.get()))
        if self.txt2.get() != "" :
            menu_array.append(self.txt2.get())
            quantity_array.append(int(self.quantity2.get()))
        if self.txt3.get() != "" :
            menu_array.append(self.txt3.get())
            quantity_array.append(int(self.quantity3.get()))
        if self.txt4.get() != "" :
            menu_array.append(self.txt4.get())
            quantity_array.append(int(self.quantity4.get()))
        if self.txt5.get() != "" :
            menu_array.append(self.txt5.get())
            quantity_array.append(int(self.quantity5.get()))
        if self.txt6.get() != "" :
            menu_array.append(self.txt6.get())
            quantity_array.append(int(self.quantity6.get()))
        if self.txt7.get() != "" :
            menu_array.append(self.txt7.get())
            quantity_array.append(int(self.quantity7.get()))
        if self.txt8.get() != "" :
            menu_array.append(self.txt8.get())
            quantity_array.append(int(self.quantity8.get()))
        if self.txt9.get() != "" :
            menu_array.append(self.txt9.get())
            quantity_array.append(int(self.quantity9.get()))
        if self.txt10.get() != "" :
            menu_array.append(self.txt10.get())
            quantity_array.append(int(self.quantity10.get()))
        pptx_auto_gen.auto_gen(self.txtdate.get(), menu_array)
        xlsx_auto_gen.auto_gen(self.txtdate.get(), menu_array, quantity_array)

    def tmp(self, parent):
        fm_input = ttk.Frame(parent, )
        parent.add(fm_input)
        # 画面の作成
        # self.main = tk.Tk()
        # self.main.geometry("350x700")

        # self.main.title("メニュー名入力したら色々できてるやつ")

        # タイトルラベル
        lbl_title = tk.Label(text='メニューと個数、日付を入力して作成ボタンを押してください')
        lbl_title.place(x=5, y=20)

        # ラベル
        lbl1 = tk.Label(text='1')
        lbl1.place(x=10, y=70)
        # テキストボックス
        # txt1 = tk.Entry(width=30)
        # txt1.place(x=30, y=70)
        # quantity1 = tk.Entry(width=2)
        # quantity1.place(x=310, y=70)
        menu_list = ['カラアゲ弁当', 'シャケカラ弁当', 'サバカラ弁当']
        self.txt1 = ttk.Combobox(values=menu_list,width=27)
        self.txt1.place(x=30, y=70)
        self.quantity1 = tk.Entry(width=2)
        self.quantity1.place(x=310, y=70)

        # ラベル
        lbl2 = tk.Label(text='2')
        lbl2.place(x=10, y=100)
        # テキストボックス
        self.txt2 = tk.Entry(width=30)
        self.txt2.place(x=30, y=100)
        self.quantity2 = tk.Entry(width=2)
        self.quantity2.place(x=310, y=100)

        # ラベル
        lbl3 = tk.Label(text='3')
        lbl3.place(x=10, y=130)
        # テキストボックス
        self.txt3 = tk.Entry(width=30)
        self.txt3.place(x=30, y=130)
        self.quantity3 = tk.Entry(width=2)
        self.quantity3.place(x=310, y=130)

        # ラベル
        lbl4 = tk.Label(text='4')
        lbl4.place(x=10, y=160)
        # テキストボックス
        self.txt4 = tk.Entry(width=30)
        self.txt4.place(x=30, y=160)
        self.quantity4 = tk.Entry(width=2)
        self.quantity4.place(x=310, y=160)

        # ラベル
        lbl5 = tk.Label(text='5')
        lbl5.place(x=10, y=190)
        # テキストボックス
        self.txt5 = tk.Entry(width=30)
        self.txt5.place(x=30, y=190)
        self.quantity5 = tk.Entry(width=2)
        self.quantity5.place(x=310, y=190)

        # ラベル
        lbl6 = tk.Label(text='6')
        lbl6.place(x=10, y=220)
        # テキストボックス
        self.txt6 = tk.Entry(width=30)
        self.txt6.place(x=30, y=220)
        self.quantity6 = tk.Entry(width=2)
        self.quantity6.place(x=310, y=220)

        # ラベル
        lbl7 = tk.Label(text='7')
        lbl7.place(x=10, y=250)
        # テキストボックス
        self.txt7 = tk.Entry(width=30)
        self.txt7.place(x=30, y=250)
        self.quantity7 = tk.Entry(width=2)
        self.quantity7.place(x=310, y=250)

        # ラベル
        lbl8 = tk.Label(text='8')
        lbl8.place(x=10, y=280)
        # テキストボックス
        self.txt8 = tk.Entry(width=30)
        self.txt8.place(x=30, y=280)
        self.quantity8 = tk.Entry(width=2)
        self.quantity8.place(x=310, y=280)

        # ラベル
        lbl9 = tk.Label(text='9')
        lbl9.place(x=10, y=310)
        # テキストボックス
        self.txt9 = tk.Entry(width=30)
        self.txt9.place(x=30, y=310)
        self.quantity9 = tk.Entry(width=2)
        self.quantity9.place(x=310, y=310)

        # ラベル
        lbl10 = tk.Label(text='10')
        lbl10.place(x=10, y=340)
        # テキストボックス
        self.txt10 = tk.Entry(width=30)
        self.txt10.place(x=30, y=340)
        self.quantity10 = tk.Entry(width=2)
        self.quantity10.place(x=310, y=340)

        # 日付
        lbldate = tk.Label(text='日付')
        lbldate.place(x=10, y=370)
        self.txtdate = tk.Entry(width=30)
        self.txtdate.place(x=50, y=370)


        # 作成ボタン
        btn1 = tk.Button(fm_input, text='作成', width=15,height=3,command = self.btn_click)
        btn1.place(x=100,y=400)



def main():
    root = tk.Tk()
    app = MainWindow(master=root)
    app.mainloop()

if __name__ == "__main__":
    show_gui = main()