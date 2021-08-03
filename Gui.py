import tkinter
import auto_gen

class Gui():


    def __init__(self):


        def btn_click():
            array = []
            array.append(txt1.get())
            array.append(txt2.get())
            array.append(txt3.get())
            array.append(txt4.get())
            array.append(txt5.get())
            array.append(txt6.get())
            array.append(txt7.get())
            array.append(txt8.get())
            array.append(txt9.get())
            array.append(txt10.get())
            print(array)

            auto_gen.auto_gen(array)

        # 画面の作成
        self.main = tkinter.Tk()
        self.main.geometry("350x700")

        # タイトルラベル
        lbl_title = tkinter.Label(text='メニューを入力して作成ボタンを押してください')
        lbl_title.place(x=5, y=20)

        # ラベル
        lbl1 = tkinter.Label(text='1')
        lbl1.place(x=10, y=70)
        # テキストボックス
        txt1 = tkinter.Entry(width=30)
        txt1.place(x=30, y=70)

        # ラベル
        lbl2 = tkinter.Label(text='2')
        lbl2.place(x=10, y=100)
        # テキストボックス
        txt2 = tkinter.Entry(width=30)
        txt2.place(x=30, y=100)

        # ラベル
        lbl3 = tkinter.Label(text='3')
        lbl3.place(x=10, y=130)
        # テキストボックス
        txt3 = tkinter.Entry(width=30)
        txt3.place(x=30, y=130)

        # ラベル
        lbl4 = tkinter.Label(text='4')
        lbl4.place(x=10, y=160)
        # テキストボックス
        txt4 = tkinter.Entry(width=30)
        txt4.place(x=30, y=160)

        # ラベル
        lbl5 = tkinter.Label(text='5')
        lbl5.place(x=10, y=190)
        # テキストボックス
        txt5 = tkinter.Entry(width=30)
        txt5.place(x=30, y=190)

        # ラベル
        lbl6 = tkinter.Label(text='6')
        lbl6.place(x=10, y=220)
        # テキストボックス
        txt6 = tkinter.Entry(width=30)
        txt6.place(x=30, y=220)

        # ラベル
        lbl7 = tkinter.Label(text='7')
        lbl7.place(x=10, y=250)
        # テキストボックス
        txt7 = tkinter.Entry(width=30)
        txt7.place(x=30, y=250)

        # ラベル
        lbl8 = tkinter.Label(text='8')
        lbl8.place(x=10, y=280)
        # テキストボックス
        txt8 = tkinter.Entry(width=30)
        txt8.place(x=30, y=280)

        # ラベル
        lbl9 = tkinter.Label(text='9')
        lbl9.place(x=10, y=310)
        # テキストボックス
        txt9 = tkinter.Entry(width=30)
        txt9.place(x=30, y=310)

        # ラベル
        lbl10 = tkinter.Label(text='10')
        lbl10.place(x=10, y=340)
        # テキストボックス
        txt10 = tkinter.Entry(width=30)
        txt10.place(x=30, y=340)

        array = []

        # 作成ボタン
        btn1 = tkinter.Button(self.main, text='送信', width=15,height=3,command = btn_click)
        btn1.place(x=60,y=370)


        # ボタンの作成
        # btn1 = tk.Button(self.main, text='フィルタ', width=15,height=3,command = self.btn_click)
        # btn1.place(x=110, y=10)

        # btn2 = tk.Button(self.main, text='生成', width=15,height=3,command = self.btn_click)
        # btn2.place(x=110, y=550)

        # btn3 = tk.Button(self.main, text='監査資料', width=15,height=3,command = self.btn_click)
        # btn3.place(x=110, y=620)

        # self.text_in = 'text'

        # 結果表示
        # self.filt = tk.Label(width=45,height=30,wraplength=400,bg="white",justify="left",anchor="nw")
        # self.filt.place(x=10,y=80)
        # self.filt[self.text_in] = self.dir_null
        self.main.mainloop()


if __name__ == "__main__":
    show_gui = Gui()