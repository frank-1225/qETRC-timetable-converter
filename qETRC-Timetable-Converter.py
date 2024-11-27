import pandas as pd
import os

def process_train_schedule(file_path, output_file):
    """
    处理单个 Excel 文件中的所有 Sheets 的列车时刻表，并将结果写入 CSV 文件
    :param file_path: 输入的 Excel 文件路径
    :param output_file: 输出的 CSV 文件路径
    """
    def clean_and_process_time(current_time, reference_time):
        """
        清理时间并根据规则填充缺失值
        :param current_time: 当前时刻值
        :param reference_time: 参考时刻值
        :return: 格式化后的时刻字符串
        """
        if isinstance(current_time, str):
            current_time = current_time.strip().replace("⠀", "").replace(" ", "")

        if pd.isna(current_time) or current_time in ["…", "--"]:
            return reference_time  # 使用参考时刻填充

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

    # 加载所有 Sheets
    sheet_data_dict = pd.read_excel(file_path, sheet_name=None, header=None)
    all_csv_data = []

    # 遍历每个 Sheet
    for sheet_name, sheet_data in sheet_data_dict.items():
        print(f"正在处理 {file_path} : {sheet_name}")

        # 询问用户该 Sheet 的方向
        while True:
            direction_input = input(f"请输入 Sheet '{sheet_name}' 的读取方向（0 为下行，1 为上行）：").strip()
            if direction_input in ['0', '1']:
                direction = int(direction_input)
                break
            else:
                print("输入无效，请重新选择方向（0 或 1）")

        row_range = range(2, sheet_data.shape[0], 2) if direction == 0 else range(sheet_data.shape[0] - 2, 0, -2)

        # 遍历每一列车次
        for col_idx in range(1, sheet_data.shape[1]):
            train_number = sheet_data.iloc[0, col_idx]
            if pd.isna(train_number):
                print(f"列 {col_idx} 没有车次编号，跳过")
                continue

            previous_departure = "00:00:00"

            # 遍历车站
            for row_idx in row_range:
                if row_idx + 1 >= sheet_data.shape[0] or row_idx < 0:
                    break

                station_name = sheet_data.iloc[row_idx, 0]
                if pd.isna(station_name):
                    continue

                # 根据方向处理到达和出发时刻
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

    # 将所有数据写入单个 CSV 文件
    csv_df = pd.DataFrame(all_csv_data)
    with open(output_file, mode='a', encoding='utf-8-sig', newline='') as f:
        csv_df.to_csv(f, index=False, header=False)

    print(f"已处理文件：{file_path} -> {output_file}")

# 遍历目录并处理所有 Excel 文件
def process_all_files_in_directory(directory_path, output_file):
    """
    遍历指定目录，处理其中的所有 Excel 文件并将结果写入同一个 CSV 文件
    :param directory_path: 输入文件所在目录
    :param output_file: 输出的 CSV 文件路径
    """
    if os.path.exists(output_file):
        os.remove(output_file)

    for file_name in os.listdir(directory_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(directory_path, file_name)
            process_train_schedule(file_path, output_file)

if __name__ == "__main__":
    input_directory = os.getcwd()  # 输入目录
    output_file = "output.csv"  # 输出文件
    process_all_files_in_directory(input_directory, output_file)
