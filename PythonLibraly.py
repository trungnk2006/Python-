import tkinter as tk
from tkinter import messagebox
import csv
import datetime
import pandas as pd


def save_to_csv():
    employee_data = [
        entries["Mã"].get(),
        entries["Tên"].get(),
        entries["Đơn vị"].get(),
        entries["Chức danh"].get(),
        entries["Ngày sinh"].get(),
        entries["Số CMND"].get(),
        entries["Nơi cấp"].get(),
        entries["Ngày cấp"].get()
    ]

    if "" in employee_data:
        messagebox.showwarning("Thiếu thông tin", "Vui lòng điền đầy đủ thông tin")
        return

    with open('employees.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(employee_data)

    messagebox.showinfo("Lưu thành công", "Thông tin nhân viên đã được lưu")


def show_birthdays_today():
    today = datetime.datetime.today().strftime('%Y-%m-%d')
    birthdays = []

    try:
        with open('employees.csv', mode='r') as file:
            reader = csv.reader(file)
            for row in reader:
                if len(row) < 5:
                    continue
                if row[4] == today:
                    birthdays.append(row)
    except FileNotFoundError:
        messagebox.showwarning("Lỗi", "Không tìm thấy file employees.csv")
        return

    if not birthdays:
        messagebox.showinfo("Không có sinh nhật", "Không có nhân viên nào có sinh nhật hôm nay")
    else:
        birthday_list = "\n".join([" - ".join(birthday) for birthday in birthdays])
        messagebox.showinfo("Sinh nhật hôm nay", f"Những nhân viên có sinh nhật hôm nay:\n{birthday_list}")


def export_to_excel():
    try:
        df = pd.read_csv('employees.csv', header=None)
        df.columns = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Số CMND", "Nơi cấp", "Ngày cấp"]
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'])
        df = df.sort_values(by='Ngày sinh', ascending=False)

        df.to_excel('employees.xlsx', index=False)
        messagebox.showinfo("Xuất thành công", "Danh sách nhân viên đã được xuất ra file employees.xlsx")
    except FileNotFoundError:
        messagebox.showwarning("Lỗi", "Không tìm thấy file employees.csv")


def main():
    root = tk.Tk()
    root.title("Quản lý thông tin nhân viên")

    labels = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Số CMND", "Nơi cấp", "Ngày cấp"]
    entries = {}
    for label in labels:
        frame = tk.Frame(root)
        frame.pack(fill='x')
        lbl = tk.Label(frame, text=label, width=15)
        lbl.pack(side='left')
        ent = tk.Entry(frame)
        ent.pack(side='left', fill='x', expand=True)
        entries[label] = ent

    save_btn = tk.Button(root, text="Lưu thông tin", command=save_to_csv)
    save_btn.pack(fill='x')

    birthday_btn = tk.Button(root, text="Sinh nhật ngày hôm nay", command=show_birthdays_today)
    birthday_btn.pack(fill='x')

    export_btn = tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel)
    export_btn.pack(fill='x')

    root.mainloop()


if __name__ == "__main__":
    main()
