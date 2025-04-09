import os
import glob
import msvcrt

# Đường dẫn tới thư mục chứa các file
folder_path = r'G:\.shortcut-targets-by-id\13cEuVXiiAR1ftfXwFQGv9VaBo4JiNv5y\2D\Đội diễn\Moho Pro\Render Cache'

# Lọc tất cả các file trong thư mục
files_to_delete = glob.glob(os.path.join(folder_path, "*"))  # Dùng "*" để khớp mọi file

total_deleted = 0
total_size = 0

# Xóa từng file trong danh sách
for file_path in files_to_delete:
    try:
        file_size = os.path.getsize(file_path)  # Lấy kích thước file trước khi xóa
        os.remove(file_path)
        total_deleted += 1
        total_size += file_size
        print(f"Đã xóa file: {file_path}")
    except Exception as e:
        print(f"Lỗi khi xóa file {file_path}: {e}")

# Hiển thị tổng kết
print(f"\nTổng số file đã xóa: {total_deleted}")
print(f"Tổng dung lượng giải phóng: {total_size / (1024 * 1024):.2f} MB")

print("\nNhấn bất kỳ phím nào để thoát...")
msvcrt.getch()
