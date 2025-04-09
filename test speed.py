import speedtest
import os
import time

def check_speed():
    st = speedtest.Speedtest()
    st.get_best_server()
    
    print("Đang kiểm tra tốc độ mạng...\n")
    time.sleep(1)  # Tạo độ trễ để hiển thị tiến trình
    
    print("Đang kiểm tra tốc độ tải xuống...")
    download_speed = st.download() / 1_000_000  # Chuyển đổi thành Mbps
    print("Đang kiểm tra tốc độ tải lên...")
    upload_speed = st.upload() / 1_000_000  # Chuyển đổi thành Mbps
    ping = st.results.ping
    
    print("\nKết quả kiểm tra:")
    print(f"Tốc độ tải xuống: {download_speed:.2f} Mbps")
    print(f"Tốc độ tải lên: {upload_speed:.2f} Mbps")
    print(f"Ping: {ping} ms")

if __name__ == "__main__":
    script_path = r"C:\Users\ADMIN\Downloads\test speed.py"
    os.chdir(os.path.dirname(script_path))
    check_speed()
    input("\nNhấn phím bất kỳ để thoát...")
    msvcrt.getch()