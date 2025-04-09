import os
import shutil
import zipfile
from pathlib import Path
from tkinter import Tk, filedialog
import msvcrt
import sys
from datetime import datetime

# Định dạng dung lượng đẹp hơn
def format_size(size_bytes):
    if size_bytes >= 1024 ** 3:
        return f"{size_bytes / (1024 ** 3):.2f} GB"
    elif size_bytes >= 1024 ** 2:
        return f"{size_bytes / (1024 ** 2):.2f} MB"
    elif size_bytes >= 1024:
        return f"{size_bytes / 1024:.2f} KB"
    else:
        return f"{size_bytes} bytes"

def count_files_and_folders(folder_path):
    total_files = 0
    total_folders = 0
    for _, dirs, files in os.walk(folder_path):
        total_files += len(files)
        total_folders += len(dirs)
    return total_files, total_folders

def print_progress_bar(current, total, filename=''):
    percent = float(current) / total if total > 0 else 0
    filename = (filename[:20] + '...') if len(filename) > 20 else filename.ljust(20)
    progress = f"\r🗂️ Đang nén: {filename} ({current}/{total}) — {percent * 100:.1f}%"
    sys.stdout.write(progress.ljust(80))
    sys.stdout.flush()

def zip_folder_with_progress(folder_path, zip_path, error_files):
    folder_path = Path(folder_path)
    all_files = list(folder_path.rglob('*'))
    files_only = [f for f in all_files if f.is_file()]
    total = len(files_only)
    count = 0

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in files_only:
            if file.suffix.lower() in ['.gsheet', '.gdoc', '.gslides', '.lnk']:
                continue
            try:
                # Kiểm tra độ dài đường dẫn
                if len(str(file)) > 260:
                    raise OSError(22, f"Đường dẫn quá dài (>260 ký tự): {file}")
                arcname = file.relative_to(folder_path)
                zipf.write(file, arcname)
                count += 1
                print_progress_bar(count, total, file.name)
            except Exception as e:
                error_msg = f"❌ Lỗi khi thêm file: {file} — {e}"
                print(f"\n{error_msg}")
                error_files.append(error_msg)  # Ghi chi tiết lỗi vào danh sách
    sys.stdout.write('\r' + ' ' * 80 + '\r')
    sys.stdout.flush()
    print()

def prepend_log_line(log_path, new_lines):
    if not log_path.exists():
        with open(log_path, 'w', encoding='utf-8') as f:
            f.writelines(line + '\n' for line in new_lines)
        return

    with open(log_path, 'r', encoding='utf-8') as f:
        old_lines = f.readlines()
    with open(log_path, 'w', encoding='utf-8') as f:
        f.writelines(line + '\n' for line in new_lines)
        f.write('\n')
        f.writelines(old_lines)

def zip_and_delete_subfolders(parent_folder, batch_size=10):
    parent_path = Path(parent_folder)
    subfolders = [f for f in parent_path.iterdir() if f.is_dir()]

    total_folders_zipped = 0
    total_files_zipped = 0
    total_folders_inside = 0
    total_saved_bytes = 0
    total_original_bytes = 0

    # Thiết lập thư mục tạm trên ổ E
    temp_dir = Path("E:/Temp")
    if not temp_dir.exists():
        temp_dir.mkdir(parents=True)
    os.environ["TMP"] = str(temp_dir)
    os.environ["TEMP"] = str(temp_dir)

    parent_name = parent_path.name
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = f"ket_qua_nen_{parent_name}_{current_time}.txt"
    log_path = parent_path / log_filename

    for batch_start in range(0, len(subfolders), batch_size):
        batch = subfolders[batch_start: batch_start + batch_size]

        for folder in batch:
            print(f"\n📦 Đang nén: {folder.name}")
            original_size = sum(f.stat().st_size for f in folder.rglob('*') if f.is_file())
            total_original_bytes += original_size

            file_count, folder_count = count_files_and_folders(folder)
            total_files_zipped += file_count
            total_folders_inside += folder_count

            zip_path = parent_path / f"{folder.name}.zip"
            if zip_path.exists():
                try:
                    zip_path.unlink()
                except Exception as e:
                    msg = f"⚠️ Không thể xóa file zip cũ: {zip_path}. Lỗi: {e}"
                    print(msg)
                    prepend_log_line(log_path, [f"❌ {folder.name} - {msg}"])
                    continue

            error_files = []
            try:
                zip_folder_with_progress(folder, zip_path, error_files)
            except OSError as e:
                if e.errno == 28:
                    msg = f"❌ Hết dung lượng trên ổ đĩa khi nén: {folder.name}. Kiểm tra ổ C hoặc E."
                else:
                    msg = f"❌ LỖI khi nén thư mục: {folder.name} - {e}"
                print(msg)
                prepend_log_line(log_path, [msg])
                continue
            except Exception as e:
                msg = f"❌ LỖI khi nén thư mục: {folder.name} - {e}"
                print(msg)
                prepend_log_line(log_path, [msg])
                continue

            compressed_size = zip_path.stat().st_size if zip_path.exists() else 0
            saved = original_size - compressed_size
            total_saved_bytes += saved

            if not error_files:
                success_lines = [
                    f"✅ Đã nén {file_count} tệp, {folder_count} thư mục con trong: {folder.name}",
                    f"📉 Giảm từ {format_size(original_size)} xuống {format_size(compressed_size)} ➜ Tiết kiệm: {format_size(saved)}",
                    f"🗑️ Đang xóa thư mục: {folder.name}"
                ]
                print(success_lines[0])
                print(success_lines[1])
                print(success_lines[2])
                shutil.rmtree(folder)
                total_folders_zipped += 1
                prepend_log_line(log_path, success_lines)
            else:
                error_msg = f"❌ Có {len(error_files)} lỗi khi nén {folder.name}. Không xóa thư mục."
                print(error_msg)
                # Ghi cả chi tiết lỗi từng file vào log
                prepend_log_line(log_path, [error_msg] + error_files)

    summary_lines = [
        "🎉 HOÀN TẤT TOÀN BỘ",
        f"📁 Đã nén & xóa: {total_folders_zipped} thư mục",
        f"📄 Tổng tệp đã nén: {total_files_zipped}",
        f"📂 Tổng thư mục con bên trong: {total_folders_inside}",
        f"💾 Tổng dung lượng giảm từ {format_size(total_original_bytes)} xuống {format_size(total_original_bytes - total_saved_bytes)} ➜ Tiết kiệm: {format_size(total_saved_bytes)}"
    ]
    prepend_log_line(log_path, summary_lines)

    print("\n🎉 HOÀN TẤT TOÀN BỘ")
    print(f"📁 Đã nén & xóa: {total_folders_zipped} thư mục")
    print(f"📄 Tổng tệp đã nén: {total_files_zipped}")
    print(f"📂 Tổng thư mục con bên trong: {total_folders_inside}")
    print(f"💾 Tổng dung lượng giảm từ {format_size(total_original_bytes)} xuống {format_size(total_original_bytes - total_saved_bytes)} ➜ Tiết kiệm: {format_size(total_saved_bytes)}")
    print(f"\n📝 Log đã lưu vào: {log_path}")
    print("\nNhấn bất kỳ phím nào để thoát...")
    msvcrt.getch()

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    print("🗂️ Vui lòng chọn thư mục cha chứa các thư mục con cần nén...")
    try:
        selected_folder = filedialog.askdirectory(title="Chọn thư mục cha")
        if selected_folder:
            zip_and_delete_subfolders(selected_folder)
        else:
            print("❌ Không có thư mục nào được chọn. Thoát.")
            print("\nNhấn bất kỳ phím nào để thoát...")
            msvcrt.getch()
    except Exception as e:
        print(f"❌ ĐÃ XẢY RA LỖI: {e}")
        print("\nNhấn bất kỳ phím nào để thoát...")
        msvcrt.getch()