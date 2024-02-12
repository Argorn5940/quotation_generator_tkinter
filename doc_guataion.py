from docxtpl import DocxTemplate

quotation_items = [
    {"detail": "pen", "unit_price": 100, "quantity": 1, "amount": 100},
    {"detail": "notebook", "unit_price": 300, "quantity": 3, "amount": 900},
    {"detail": "luck", "unit_price": 500, "quantity": 3, "amount": 1500},]

doc = DocxTemplate("quotation_template.docx")
doc.render({"company":"足立株式会社",
            "name":"齋藤 広宣",
            "manager":"齋藤 美紀",
            "quotation_items": quotation_items,
            "subtotal":2500,
            "salestax":"8%",
            "total":2500})
doc.save("new_quotation.docx")