import tkinter
import sells_report_auto_gen

class Gui():


    def __init__(self):


        def btn_click1():
            # menu_array = []
            # quantity_array = []
            # if txt1.get() != "" :
            #     menu_array.append(txt1.get())
            #     quantity_array.append(int(quantity1.get()))
            # if txt2.get() != "" :
            #     menu_array.append(txt2.get())
            #     quantity_array.append(int(quantity2.get()))
            # if txt3.get() != "" :
            #     menu_array.append(txt3.get())
            #     quantity_array.append(int(quantity3.get()))
            # if txt4.get() != "" :
            #     menu_array.append(txt4.get())
            #     quantity_array.append(int(quantity4.get()))
            # if txt5.get() != "" :
            #     menu_array.append(txt5.get())
            #     quantity_array.append(int(quantity5.get()))
            # if txt6.get() != "" :
            #     menu_array.append(txt6.get())
            #     quantity_array.append(int(quantity6.get()))
            # if txt7.get() != "" :
            #     menu_array.append(txt7.get())
            #     quantity_array.append(int(quantity7.get()))
            # if txt8.get() != "" :
            #     menu_array.append(txt8.get())
            #     quantity_array.append(int(quantity8.get()))
            # if txt9.get() != "" :
            #     menu_array.append(txt9.get())
            #     quantity_array.append(int(quantity9.get()))
            # if txt10.get() != "" :
            #     menu_array.append(txt10.get())
            #     quantity_array.append(int(quantity10.get()))
            # pptx_auto_gen.auto_gen(date_str.get(), menu_array)
            # xlsx_auto_gen.auto_gen(date_str.get(), menu_array, quantity_array)
            sells_report_auto_gen.get_todays_menu(date_str)



        # 画面の作成
        self.main = tkinter.Tk()
        self.main.geometry("350x700")

        self.main.title("押したら一瞬で売上報告書できてるやつ")

        # タイトルラベル
        lbl_title = tkinter.Label(text='日付を入力してください')
        lbl_title.place(x=5, y=20)


        # 日付
        lbldate = tkinter.Label(text='日付')
        lbldate.place(x=10, y=70)
        date_str = tkinter.Entry(width=30)
        date_str.place(x=50, y=70)

        # メニュー表示ボタン
        btn1 = tkinter.Button(self.main, text='メニュー表示', width=15,height=3,command = btn_click1)
        btn1.place(x=100,y=100)

        self.main.mainloop()

if __name__ == "__main__":
    show_gui = Gui()