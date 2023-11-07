import tkinter as tk
from openpyxl import Workbook
from tkinter import messagebox
from tkcalendar import Calendar

class ExpenseTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("가계부 어플리케이션")
        self.transactions = []
        self.create_ui()

    def create_ui(self):
        # 날짜 선택
        date_label = tk.Label(self.root, text="날짜:")
        date_label.pack()
        self.date_cal = Calendar(self.root, selectmode="day", year=2023, month=1, day=1)
        self.date_cal.pack()

        # 카테고리 선택
        category_label = tk.Label(self.root, text="카테고리:")
        category_label.pack()
        self.category_var = tk.StringVar()
        self.category_var.set("식비")
        category_menu = tk.OptionMenu(self.root, self.category_var, "식비", "교통비", "고정지출", "변동지출", "여행")
        category_menu.pack()

        # 금액 입력
        amount_label = tk.Label(self.root, text="금액:")
        amount_label.pack()
        self.amount_entry = tk.Entry(self.root)
        self.amount_entry.pack()

        # 설명 입력
        description_label = tk.Label(self.root, text="설명:")
        description_label.pack()
        self.description_entry = tk.Entry(self.root)
        self.description_entry.pack()

        # 저장 버튼
        save_button = tk.Button(self.root, text="저장", command=self.save_transaction)
        save_button.pack()

        # 거래 내역 보기 버튼
        show_transactions_button = tk.Button(self.root, text="거래 내역 보기", command=self.show_transactions)
        show_transactions_button.pack()

    def save_transaction(self):
        date = self.date_cal.get_date()
        category = self.category_var.get()
        amount = self.amount_entry.get()
        description = self.description_entry.get()

        if date and category and amount:
            try:
                amount = float(amount)
                self.transactions.append([date, category, amount, description])
                messagebox.showinfo("성공", "거래가 저장되었습니다.")
                self.clear_inputs()
            except ValueError:
                messagebox.showerror("오류", "유효한 금액을 입력하세요.")
        else:
            messagebox.showerror("오류", "날짜, 카테고리 및 금액은 필수 입력 항목입니다.")

    def clear_inputs(self):
        self.date_cal.set_date("2023-01-01")
        self.amount_entry.delete(0, tk.END)
        self.description_entry.delete(0, tk.END)

    def show_transactions(self):
        if not self.transactions:
            messagebox.showinfo("거래 내역", "거래 내역이 없습니다.")
            return

        # 엑셀 파일에 거래 내역 저장
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["날짜", "카테고리", "금액", "설명"])
        for transaction in self.transactions:
            sheet.append(transaction)

        # 엑셀 파일 저장
        filename = "가계부.xlsx"
        workbook.save(filename)
        messagebox.showinfo("거래 내역", f"거래 내역이 '{filename}' 파일로 저장되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExpenseTrackerApp(root)
    root.mainloop()
