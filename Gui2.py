import tkinter
import sells_report_auto_gen
import pandas

class Gui():


    def __init__(self):
        def btn_click1():
            sells_report_auto_gen.get_todays_menu(date_str.get())
            frame2.tkraise()
    
 

        # 画面の作成
        self.main = tkinter.Tk()
        self.main.geometry("350x700")

        self.main.title("押したら一瞬で売上報告書できてるやつ")

        self.main.grid_rowconfigure(0, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        # 画面1フレーム作成
        frame1 = tkinter.Frame()
        frame1.grid(row=0, column=0, sticky="nsew", pady=20)

        # 操作指示コメント
        lbl_title = tkinter.Label(frame1,text='日付を入力してください')
        lbl_title.place(x=5, y=20)

        # 日付入力欄
        lbldate = tkinter.Label(frame1, text='日付')
        lbldate.place(x=10, y=70)
        date_str = tkinter.Entry(frame1, width=30)
        date_str.place(x=50, y=70)

        # メニュー表示ボタン
        btn1 = tkinter.Button(frame1, text='メニュー表示', width=15,height=3,command = btn_click1)
        btn1.place(x=100,y=100)

        # 画面2フレーム作成
        frame2 = tkinter.Frame(self.main)
        frame2.grid(row=0, column=0, sticky="nsew", pady=20)
        lbl_title2 = tkinter.Label(frame2, text='切り替えテスト')
        lbl_title2.place(x=5, y=20)

        frame1.tkraise()

        self.main.mainloop()

if __name__ == "__main__":
    show_gui = Gui()