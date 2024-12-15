import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def read_excel(file_path):
    # Đọc file Excel
    try:
        # Đọc tất cả sheet trong file Excel
        df = pd.read_excel(file_path, sheet_name=None)
        return df
    except Exception as e:
        print(f"Đã có lỗi khi đọc file Excel: {e}")
        return None

def find_duplicate_questions(df, output_file):
    # Duyệt qua tất cả các sheet và tìm các giá trị trùng trong cột 'question'
    questions_dict = {}  # Lưu trữ tất cả các giá trị 'question' và thông tin sheet + dòng

    if df:
        # Duyệt từng sheet trong DataFrame
        for sheet_name, sheet_data in df.items():
            # Kiểm tra xem cột 'question' có tồn tại trong sheet hay không
            if 'question' in sheet_data.columns:
                questions = sheet_data['question'].dropna()  # Lấy tất cả giá trị không phải NaN
                for i, question in enumerate(questions):
                    # Nếu giá trị question đã tồn tại trong dictionary, thêm thông tin dòng và sheet
                    if question in questions_dict:
                        questions_dict[question].append((sheet_name, i + 1))  # Dòng bắt đầu từ 1
                    else:
                        questions_dict[question] = [(sheet_name, i + 1)]

        # Ghi kết quả vào file
        with open(output_file, "w", encoding="utf-8") as file:
            file.write("Các giá trị trùng trong cột 'question':\n")
            for question, occurrences in questions_dict.items():
                if len(occurrences) > 1:  # Chỉ ghi khi có sự trùng lặp
                    file.write(f"\nGiá trị: {question}\n")
                    for sheet_name, row_num in occurrences:
                        file.write(f"- Sheet: {sheet_name}, Dòng: {row_num}\n")

        print(f"Kết quả đã được ghi vào file: {output_file}")
    else:
        print("Không có dữ liệu để kiểm tra.")

def select_file():
    # Mở cửa sổ chọn file
    Tk().withdraw()  # Ẩn cửa sổ chính của tkinter
    file_path = askopenfilename(title="Chọn file Excel", filetypes=[("Excel Files", "*.xls;*.xlsx")])
    return file_path

def main():
    # Mở cửa sổ chọn file Excel
    file_path = select_file()

    if file_path:
        # Đọc file Excel
        excel_data = read_excel(file_path)

        # Đường dẫn file txt đầu ra
        output_file = "duplicate_questions.txt"

        # Tìm và ghi các giá trị trùng trong cột 'question' vào file txt
        find_duplicate_questions(excel_data, output_file)
    else:
        print("Không có file nào được chọn.")

if __name__ == "__main__":
    main()
