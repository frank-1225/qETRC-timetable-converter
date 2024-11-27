import pandas as pd
import os
from tkinter import Tk, filedialog, ttk, StringVar, messagebox, Button, Label, Entry, Frame, Canvas, Scrollbar

def process_train_schedule(file_path, sheet_config, output_file):
    def clean_and_process_time(current_time, reference_time):
        if isinstance(current_time, str):
            current_time = current_time.strip().replace("⠀", "").replace(" ", "")
        if pd.isna(current_time) or current_time in ["…", "--"]:
            return reference_time
        if isinstance(current_time, str):
            if ":" in current_time:
                if len(current_time.split(":")[1]) == 4:
                    parts = current_time.split(":")
                    hour = parts[0].zfill(2)
                    minutes = parts[1][:2]
                    seconds = parts[1][2:]
                    return f"{hour}:{minutes}:{seconds}"
                else:
                    return f"{current_time}:00"
            elif len(current_time) == 2:
                previous_hour = int(reference_time.split(":")[0])
                return f"{previous_hour:02}:{current_time}:00"
            elif len(current_time) == 4:
                previous_hour = int(reference_time.split(":")[0])
                minutes, seconds = current_time[:2], current_time[2:]
                return f"{previous_hour:02}:{minutes}:{seconds}"
        return str(current_time)

    sheet_data_dict = pd.read_excel(file_path, sheet_name=None, header=None)
    all_csv_data = []

    for sheet_name, direction in sheet_config.items():
        sheet_data = sheet_data_dict[sheet_name]
        row_range = range(2, sheet_data.shape[0], 2) if direction == 0 else range(sheet_data.shape[0] - 2, 0, -2)

        for col_idx in range(1, sheet_data.shape[1]):
            train_number = sheet_data.iloc[0, col_idx]
            if pd.isna(train_number):
                continue

            previous_departure = "00:00:00"

            for row_idx in row_range:
                if row_idx + 1 >= sheet_data.shape[0] or row_idx < 0:
                    break

                station_name = sheet_data.iloc[row_idx, 0]
                if pd.isna(station_name):
                    continue

                if direction == 1:
                    raw_arrival_time = sheet_data.iloc[row_idx + 1, col_idx]
                    raw_departure_time = sheet_data.iloc[row_idx, col_idx]
                else:
                    raw_arrival_time = sheet_data.iloc[row_idx, col_idx]
                    raw_departure_time = sheet_data.iloc[row_idx + 1, col_idx]

                if (pd.isna(raw_arrival_time) or raw_arrival_time in ["…", "--"]) and \
                   (pd.isna(raw_departure_time) or raw_departure_time in ["…", "--"]):
                    continue

                arrival_time = clean_and_process_time(raw_arrival_time, previous_departure)
                departure_time = clean_and_process_time(raw_departure_time, arrival_time)

                if pd.isna(raw_arrival_time) or raw_arrival_time in ["…", "--"]:
                    arrival_time = departure_time

                previous_departure = departure_time
                all_csv_data.append([train_number, station_name, arrival_time, departure_time])

    csv_df = pd.DataFrame(all_csv_data)
    with open(output_file, mode='a', encoding='utf-8-sig', newline='') as f:
        csv_df.to_csv(f, index=False, header=False)

def browse_input_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    if files:
        input_file_paths.extend(files)
        input_file_paths_text.set("; ".join(input_file_paths))  # 将选中的文件路径更新到文本框
        load_sheets()


def browse_output_file():
    output_file_path.set(filedialog.asksaveasfilename(filetypes=[("CSV Files", "*.csv")], defaultextension=".csv"))


def load_sheets():
    for widget in sheets_frame.winfo_children():
        widget.destroy()  # 清空之前的内容

    for file_path in input_file_paths:
        if not os.path.isfile(file_path):
            messagebox.showerror("错误", "请选择有效的输入文件")
            return

        sheet_names = pd.ExcelFile(file_path).sheet_names
        for sheet_name in sheet_names:
            row = len(sheet_configs)
            sheet_configs.append((file_path, sheet_name))
            Label(sheets_frame, text=os.path.basename(file_path), anchor="w", width=20).grid(row=row, column=0, sticky="w")
            Label(sheets_frame, text=sheet_name, anchor="w", width=30).grid(row=row, column=1, sticky="w")
            direction_var = StringVar(value="0 - 下行")
            direction_dropdown = ttk.Combobox(sheets_frame, textvariable=direction_var, values=["0 - 下行", "1 - 上行"], width=10)
            direction_dropdown.grid(row=row, column=2, padx=10)
            sheet_direction_vars.append(direction_var)


def process_files():
    if not input_file_paths:
        messagebox.showerror("错误", "请选择输入文件")
        return

    if not output_file_path.get():
        messagebox.showerror("错误", "请选择输出文件路径")
        return

    # 将 sheet 的配置保存为字典，键为 (file_path, sheet_name)，值为方向
    final_sheet_config = {(file_path, sheet_name): int(var.get().split(" ")[0]) 
                          for (file_path, sheet_name), var in zip(sheet_configs, sheet_direction_vars)}

    for (file_path, sheet_name), direction in final_sheet_config.items():
        process_train_schedule(file_path, {sheet_name: direction}, output_file_path.get())

    messagebox.showinfo("完成", f"文件已处理完成并写入文件末尾！\n输出路径：{output_file_path.get()}")

def set_all_directions(direction):
    for direction_var in sheet_direction_vars:
        direction_var.set(str(direction) + " - " + ("下行" if direction == 0 else "上行"))

# 初始化 GUI 界面
root = Tk()
root.title("qETRC列车时刻表转换工具")

# 输入文件路径
Label(root, text="输入文件:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
input_file_paths_text = StringVar()
Entry(root, textvariable=input_file_paths_text, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
Button(root, text="选择", command=browse_input_files).grid(row=0, column=2, padx=5, pady=5)

# 输出文件路径
Label(root, text="输出文件:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
output_file_path = StringVar()
Entry(root, textvariable=output_file_path, width=50).grid(row=1, column=1, padx=5, pady=5)
Button(root, text="选择", command=browse_output_file).grid(row=1, column=2, padx=5, pady=5)

# Sheets 配置表格 + 滚动条
Label(root, text="工作表配置:").grid(row=2, column=0, sticky="nw", padx=5, pady=5)

canvas = Canvas(root)
canvas.grid(row=2, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.grid(row=2, column=3, sticky="ns", padx=5, pady=5)

canvas.configure(yscrollcommand=scrollbar.set)
sheets_frame = Frame(canvas)
canvas.create_window((0, 0), window=sheets_frame, anchor="nw")

sheets_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# 设置按钮
Button(root, text="全部设为下行", command=lambda: set_all_directions(0), width=15).grid(row=3, column=0, pady=20)
Button(root, text="全部设为上行", command=lambda: set_all_directions(1), width=15).grid(row=3, column=1, pady=20)
Button(root, text="开始处理", command=process_files, width=15).grid(row=3, column=2, pady=20)

input_file_paths = []
sheet_configs = []
sheet_direction_vars = []

root.mainloop()
