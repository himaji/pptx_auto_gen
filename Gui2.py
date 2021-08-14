import tkinter as tk
import tkinter.ttk as ttk
import sells_report_auto_gen
import pandas as pd
import datetime as dt

class MainWindow(tk.Frame):
    def __init__(self, master=None, parent=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("800x700")
        self.master.title("売上報告書できるやつ")
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

        self.set_data()
        self.create_widgets()

    def set_data(self):
        self.data = pd.read_excel("../output/sales_management.xlsx")
        df = pd.read_excel("../output/sales_management.xlsx")
        flg = df["日付"].astype("str").str.isdigit()
        from_serial = pd.to_timedelta(df.loc[flg, "日付"].astype("int"), unit="D") + pd.to_datetime("1900/1/1")
        from_string = pd.to_datetime(df.loc[~flg, "日付"])
        df["日付"] = pd.concat([from_serial, from_string])
        self.colname_list = ["日付", "弁当名等", "搬入個数", "販売個数"]  # 結果に表示させる列名
        self.width_list = [100, 200]
        self.search_col = "日付"  # 検索キーワードの入力されている列名

    def create_widgets(self):
        self.pw_main = ttk.PanedWindow(self.master, orient="vertical")
        self.pw_main.pack(expand=True, fill=tk.BOTH, side="left")
        self.pw_top = ttk.PanedWindow(self.pw_main, orient="horizontal", height=25)
        self.pw_main.add(self.pw_top)
        self.pw_bottom = ttk.PanedWindow(self.pw_main, orient="vertical")
        self.pw_main.add(self.pw_bottom)
        self.creat_input_frame(self.pw_top)
        self.create_tree(self.pw_bottom)

    def creat_input_frame(self, parent):
        fm_input = ttk.Frame(parent, )
        parent.add(fm_input)
        lbl_keyword = ttk.Label(fm_input, text="キーワード", width=7)
        lbl_keyword.grid(row=1, column=1, padx=2, pady=2)
        self.keyword = tk.StringVar()
        ent_keyword = ttk.Entry(fm_input, justify="left", textvariable=self.keyword)
        ent_keyword.grid(row=1, column=2, padx=2, pady=2)
        ent_keyword.bind("<Return>", self.search)

    def create_tree(self, parent):
        self.result_text = tk.StringVar()
        lbl_result = ttk.Label(parent, textvariable=self.result_text)
        parent.add(lbl_result)
        self.tree = ttk.Treeview(parent)
        self.tree["column"] = self.colname_list
        self.tree["show"] = "headings"
        self.tree.bind("<Double-1>", self.onDouble)
        for i, (colname, width) in enumerate(zip(self.colname_list, self.width_list)):
            self.tree.heading(i, text=colname)
            self.tree.column(i, width=width)
        parent.add(self.tree)

    def search(self, event=None):
        keyword = dt.datetime.strptime(self.keyword.get(), '%Y-%m-%d')
        print(keyword)
        result = self.data[self.data['日付'] == keyword]
        self.update_tree_by_search_result(result)

    def update_tree_by_search_result(self, result):
        self.tree.delete(*self.tree.get_children())
        self.result_text.set(f"検索結果：{len(result)}")
        for _, row in result.iterrows():
            self.tree.insert("", "end", values=row[self.colname_list].to_list())

    def onDouble(self, event):
        for item in self.tree.selection():
            print(self.tree.item(item)["values"])

def main():
    root = tk.Tk()
    app = MainWindow(master=root)
    app.mainloop()

if __name__ == "__main__":
    show_gui = main()