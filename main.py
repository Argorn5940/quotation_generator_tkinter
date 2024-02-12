import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

quotation_list = []
def add_button():
  """
  add_button 関数は、新しい項目を引用リストに追加し、ツリー ビュー ウィジェットに挿入します。
  """
  detail = detail_entry.get()
  unit_price_value = int(unit_price_entry.get())
  qty = int(quantity_spinbox.get())
  line_total = qty*unit_price_value
  quotation_item = [detail, unit_price_value, qty, line_total]
  tree.insert('',0, values=quotation_item)
  clear_item()
  quotation_list.append(quotation_item)

  """
  関数 `clear_item()` は、`detail_entry`、`unit_price_entry`、および `quantity_spinbox`
  フィールドの値をクリアし、`quantity_spinbox` の値を "1" に設定します。
  """
def clear_item():
  detail_entry.delete(0,tkinter.END)
  unit_price_entry.delete(0,tkinter.END)
  quantity_spinbox.delete(0,tkinter.END) # まず現在の値を消去します
  quantity_spinbox.insert(0,"1") # 次に新しい値を挿入します

"""
関数 `new_quotation()` は、見積フォームの入力フィールドとツリー ビューをクリアします。
"""
def new_quotation():
  company_name_entry.delete(0,tkinter.END)
  first_name_entry.delete(0, tkinter.END)
  last_name_entry.delete(0, tkinter.END)
  clear_item()
  tree.delete(*tree.get_children())

  """
  関数 `generate_quote()` は、テンプレート、入力データ、および計算を使用して見積書を生成します。
  """
def generate_quotation():
    doc = DocxTemplate("quotation_template.docx")
    company = company_name_entry.get()
    name = first_name_entry.get() + " " + last_name_entry.get()
    manager = manager_first_name_entry.get() + " " + manager_last_name_entry.get()
    # quotation_listから新しい辞書リストを作成する
    transformed_quotation_list = [
        {'detail': item[0], 'unit_price': item[1], 'quantity': item[2], 'amount': item[3]}
        for item in quotation_list
    ]
    subtotal = sum(item['amount'] for item in transformed_quotation_list)
    salestax_rate = 0.1
    salestax_amount = subtotal * salestax_rate
    total = subtotal + salestax_amount
    doc.render({
        "company": company,
        "name": name,
        "manager": manager,
        "quotation_items": transformed_quotation_list,  # 変換したリストを使用
        "subtotal": subtotal,
        "salestax": "{:.2f}%".format(salestax_rate*100),
        "total": total
    })
    # ドキュメントの名前を変数に設定し、保存メソッドに直接渡します。
    file_name = "new_quotation_" + name.replace(' ', '_') + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(file_name)
    messagebox.showinfo("請求書作成", "見積書作成が完了しました！")

# このコードは、Python の tkinter ライブラリを使用してグラフィカル ユーザー インターフェイス (GUI)
# を作成しています。ウィンドウを作成し、ラベル、入力フィールド、ボタン、ツリービュー ウィジェットなどのさまざまなウィジェットを追加します。
window = tkinter.Tk()
window.title("quotation Generator Form")

frame = tkinter.Frame(window)
frame.pack(padx=60, pady=20)
company_name_label = tkinter.Label(frame, text="会社名")
company_name_label.grid(row=0,column=0)
first_name_label = tkinter.Label(frame, text="顧客性")
first_name_label.grid(row=0,column=1)
last_name_label = tkinter.Label(frame, text="顧客名")
last_name_label.grid(row=0, column=2)
manager_first_name = tkinter.Label(frame, text="担当者性")
manager_first_name.grid(row=0, column=3)
manager_last_name = tkinter.Label(frame, text="担当者名")
manager_last_name.grid(row=0, column=4)
detail_label = tkinter.Label(frame, text="詳細")
detail_label.grid(row=2, column=0)
unit_price_label = tkinter.Label(frame, text="単価")
unit_price_label.grid(row=2, column=1)
quantity_label = tkinter.Label(frame, text="数量")
quantity_label.grid(row=2, column=2)
company_name_entry = tkinter.Entry(frame, width=30,)
first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
manager_first_name_entry = tkinter.Entry(frame)
manager_last_name_entry = tkinter.Entry(frame)
company_name_entry.grid(row=1, column=0)
first_name_entry.grid(row=1, column=1)
last_name_entry.grid(row=1, column=2, padx=(1, 1))
manager_first_name_entry.grid(row=1, column=3, padx=(0, 15))
manager_last_name_entry.grid(row=1, column=4)
detail_entry = tkinter.Entry(frame, width=40)
detail_entry.grid(row=3, column=0)
unit_price_entry = tkinter.Entry(frame)
unit_price_entry.grid(row=3, column=1)
quantity_spinbox = tkinter.Spinbox(frame, from_=0, to=100)
quantity_spinbox.grid(row=3, column=2)
add_buttom = tkinter.Button(frame, width=10, text="追加", command = add_button)
add_buttom.grid(row=3, column=3)
columns = ('detail', 'unit_price', 'quantity', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.grid(row=4, column=0, columnspan=3, padx=20, pady=10)
tree.heading('detail', text='詳細')
tree.heading('unit_price', text='単価')
tree.heading('quantity', text='数量')
tree.heading('total', text='合計')

# このコードは、tkinter ライブラリを使用してボタン ウィジェットを作成しています。ボタンには「見積書作成」（英語で「Generate
# Quotation」を意味します）というラベルが付いており、クリックすると「generate_quote()」関数が実行されます。次に、ボタンは GUI ウィンドウの行 5、列 0 に配置され、3
# 列にまたがります。 「sticky」パラメータは「news」に設定されています。これは、ボタンが南北方向と東西方向の両方の利用可能なスペースを埋めるように拡張されることを意味します。 `padx`
# および `pady` パラメータは、ボタンの周囲にパディングを追加するために使用されます。

save_quotation_button = tkinter.Button(frame, text="見積書作成", command=generate_quotation)
save_quotation_button.grid(row=5, column=0, columnspan=3, sticky="news", padx=20, pady=5)

# コード `new_quotation_button = tkinter.Button(frame, text="新見積書", command=new_quotation)`
# は、ラベル「新見積書」（英語で「New Quotation」を意味します）を持つボタン ウィジェットを作成し、
# それを変数「new_quote_button」。 `command`
# パラメータは、ボタンがクリックされたときに実行される関数 `new_quote` を指定します。
new_quotation_button = tkinter.Button(frame, text="新見積書", command=new_quotation)
new_quotation_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
window.mainloop()