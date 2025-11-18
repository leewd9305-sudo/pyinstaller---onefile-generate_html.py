from tkinter import Tk, filedialog, messagebox

def select_excel_file():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="제작가이드 엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        raise FileNotFoundError("엑셀 파일이 선택되지 않았습니다!")

    return file_path


def show_info(msg):
    messagebox.showinfo("알림", msg)


def show_error(msg):
    messagebox.showerror("오류 발생", msg)
