import tkinter
import pptx_auto_gen
import xlsx_auto_gen

class Gui():


    def __init__(self):


        def btn_click():
            menu_array = []
            quantity_array = []
            if txt1.get() != "" :
                menu_array.append(txt1.get())
                quantity_array.append(int(quantity1.get()))
            if txt2.get() != "" :
                menu_array.append(txt2.get())
                quantity_array.append(int(quantity2.get()))
            if txt3.get() != "" :
                menu_array.append(txt3.get())
                quantity_array.append(int(quantity3.get()))
            if txt4.get() != "" :
                menu_array.append(txt4.get())
                quantity_array.append(int(quantity4.get()))
            if txt5.get() != "" :
                menu_array.append(txt5.get())
                quantity_array.append(int(quantity5.get()))
            if txt6.get() != "" :
                menu_array.append(txt6.get())
                quantity_array.append(int(quantity6.get()))
            if txt7.get() != "" :
                menu_array.append(txt7.get())
                quantity_array.append(int(quantity7.get()))
            if txt8.get() != "" :
                menu_array.append(txt8.get())
                quantity_array.append(int(quantity8.get()))
            if txt9.get() != "" :
                menu_array.append(txt9.get())
                quantity_array.append(int(quantity9.get()))
            if txt10.get() != "" :
                menu_array.append(txt10.get())
                quantity_array.append(int(quantity10.get()))
            pptx_auto_gen.auto_gen(txtdate.get(), menu_array)
            xlsx_auto_gen.auto_gen(txtdate.get(), menu_array, quantity_array)



        # 画面の作成
        self.main = tkinter.Tk()
        self.main.geometry("350x700")

        self.main.title("メニュー名入力したら色々できてるやつ")

        # タイトルラベル
        lbl_title = tkinter.Label(text='メニューと個数、日付を入力して作成ボタンを押してください')
        lbl_title.place(x=5, y=20)

        # ラベル
        lbl1 = tkinter.Label(text='1')
        lbl1.place(x=10, y=70)
        # テキストボックス
        txt1 = tkinter.Entry(width=30)
        txt1.place(x=30, y=70)
        quantity1 = tkinter.Entry(width=2)
        quantity1.place(x=310, y=70)

        # ラベル
        lbl2 = tkinter.Label(text='2')
        lbl2.place(x=10, y=100)
        # テキストボックス
        txt2 = tkinter.Entry(width=30)
        txt2.place(x=30, y=100)
        quantity2 = tkinter.Entry(width=2)
        quantity2.place(x=310, y=100)

        # ラベル
        lbl3 = tkinter.Label(text='3')
        lbl3.place(x=10, y=130)
        # テキストボックス
        txt3 = tkinter.Entry(width=30)
        txt3.place(x=30, y=130)
        quantity3 = tkinter.Entry(width=2)
        quantity3.place(x=310, y=130)

        # ラベル
        lbl4 = tkinter.Label(text='4')
        lbl4.place(x=10, y=160)
        # テキストボックス
        txt4 = tkinter.Entry(width=30)
        txt4.place(x=30, y=160)
        quantity4 = tkinter.Entry(width=2)
        quantity4.place(x=310, y=160)

        # ラベル
        lbl5 = tkinter.Label(text='5')
        lbl5.place(x=10, y=190)
        # テキストボックス
        txt5 = tkinter.Entry(width=30)
        txt5.place(x=30, y=190)
        quantity5 = tkinter.Entry(width=2)
        quantity5.place(x=310, y=190)

        # ラベル
        lbl6 = tkinter.Label(text='6')
        lbl6.place(x=10, y=220)
        # テキストボックス
        txt6 = tkinter.Entry(width=30)
        txt6.place(x=30, y=220)
        quantity6 = tkinter.Entry(width=2)
        quantity6.place(x=310, y=220)

        # ラベル
        lbl7 = tkinter.Label(text='7')
        lbl7.place(x=10, y=250)
        # テキストボックス
        txt7 = tkinter.Entry(width=30)
        txt7.place(x=30, y=250)
        quantity7 = tkinter.Entry(width=2)
        quantity7.place(x=310, y=250)

        # ラベル
        lbl8 = tkinter.Label(text='8')
        lbl8.place(x=10, y=280)
        # テキストボックス
        txt8 = tkinter.Entry(width=30)
        txt8.place(x=30, y=280)
        quantity8 = tkinter.Entry(width=2)
        quantity8.place(x=310, y=280)

        # ラベル
        lbl9 = tkinter.Label(text='9')
        lbl9.place(x=10, y=310)
        # テキストボックス
        txt9 = tkinter.Entry(width=30)
        txt9.place(x=30, y=310)
        quantity9 = tkinter.Entry(width=2)
        quantity9.place(x=310, y=310)

        # ラベル
        lbl10 = tkinter.Label(text='10')
        lbl10.place(x=10, y=340)
        # テキストボックス
        txt10 = tkinter.Entry(width=30)
        txt10.place(x=30, y=340)
        quantity10 = tkinter.Entry(width=2)
        quantity10.place(x=310, y=340)

        # 日付
        lbldate = tkinter.Label(text='日付')
        lbldate.place(x=10, y=370)
        txtdate = tkinter.Entry(width=30)
        txtdate.place(x=50, y=370)


        # 作成ボタン
        btn1 = tkinter.Button(self.main, text='作成', width=15,height=3,command = btn_click)
        btn1.place(x=100,y=400)

        self.main.mainloop()


if __name__ == "__main__":
    show_gui = Gui()