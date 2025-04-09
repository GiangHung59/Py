import os
import shutil
import ctypes
import msvcrt

def empty_temp():
    temp_path = os.getenv('TEMP')
    total_deleted = 0
    total_size = 0
    
    if temp_path and os.path.exists(temp_path):
        try:
            for filename in os.listdir(temp_path):
                file_path = os.path.join(temp_path, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        total_size += file_size
                        total_deleted += 1
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path, ignore_errors=True)
                        total_deleted += 1
                except PermissionError:
                    print(f'Bỏ qua: {file_path} (đang được sử dụng)')
                except Exception as e:
                    if "[WinError 2]" in str(e):
                        print(f'Không tìm thấy tập tin: {file_path}')
                    else:
                        print(f'Lỗi khi xóa {file_path}: {e}')
        except Exception as e:
            print(f'Lỗi khi truy cập {temp_path}: {e}')
    else:
        print('Thư mục TEMP không tồn tại.')
    
    print(f'Đã xóa {total_deleted} mục, giải phóng khoảng {total_size / (1024 * 1024):.2f} MB.')

def empty_recycle_bin():
    try:
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 1)
        print("Đã dọn sạch thùng rác.")
    except Exception as e:
        print(f'Lỗi khi xóa thùng rác: {e}')

if __name__ == "__main__":
    empty_temp()
    empty_recycle_bin()
    print("Nhấn bất kỳ phím nào để thoát...")
    msvcrt.getch()
