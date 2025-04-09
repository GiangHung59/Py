import ctypes
import time
import os
import msvcrt

def get_recycle_bin_size():
    try:
        recycle_bin_path = os.path.expandvars('%SYSTEMDRIVE%\\$Recycle.Bin')
        total_size = 0
        total_files = 0
        
        for root, dirs, files in os.walk(recycle_bin_path):
            total_files += len(files)
            total_size += sum(os.path.getsize(os.path.join(root, f)) for f in files)
        
        return total_files, total_size / (1024 * 1024)  # Chuyển đổi sang MB
    except Exception as e:
        print(f'Lỗi khi lấy thông tin thùng rác: {e}')
        return 0, 0

def empty_recycle_bin():
    try:
        total_files, total_size = get_recycle_bin_size()
        start_time = time.time()
        
        # Windows API để dọn sạch thùng rác
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 1)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        print(f"Đã xóa {total_files} tệp, giải phóng {total_size:.2f} MB.")
        print(f"Thời gian thực hiện: {elapsed_time:.2f} giây.")
    except Exception as e:
        print(f'Lỗi khi xóa thùng rác: {e}')

if __name__ == "__main__":
    empty_recycle_bin()
    print("Nhấn bất kỳ phím nào để thoát...")
    msvcrt.getch()
