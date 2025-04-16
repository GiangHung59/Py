# Đảm bảo mã hóa UTF-8 cho đầu ra
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Tải assembly để hiển thị hộp thoại chọn thư mục
Add-Type -AssemblyName System.Windows.Forms

# Hiển thị hộp thoại chọn thư mục
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Chon thu muc chua cac thu muc can nen"
$folderBrowser.ShowNewFolderButton = $false

# Nếu người dùng chọn OK, lấy đường dẫn
if ($folderBrowser.ShowDialog() -eq "OK") {
    $sourcePath = $folderBrowser.SelectedPath
} else {
    Write-Warning "Khong chon thu muc. Thoat script."
    exit
}

# Đường dẫn nơi lưu log lỗi
$logPath = "E:\Error_Log.txt"

# Kiểm tra xem WinRAR có được cài đặt không
$winrarPath = "C:\Program Files\WinRAR\WinRAR.exe"
if (-not (Test-Path $winrarPath)) {
    $errorMsg = "$(Get-Date) - Loi: WinRAR khong duoc cai dat tai $winrarPath"
    Add-Content -Path $logPath -Value $errorMsg -Encoding UTF8
    Write-Warning $errorMsg
    exit
}

# Kiểm tra xem đường dẫn nguồn có tồn tại không
if (-not (Test-Path $sourcePath)) {
    $errorMsg = "$(Get-Date) - Loi: Duong dan $sourcePath khong ton tai"
    Add-Content -Path $logPath -Value $errorMsg -Encoding UTF8
    Write-Warning $errorMsg
    exit
}

# Lấy danh sách các thư mục trong đường dẫn nguồn
$folders = Get-ChildItem -Path $sourcePath -Directory -ErrorAction Stop

if ($folders.Count -eq 0) {
    $warningMsg = "$(Get-Date) - Canh bao: Khong tim thay thu muc nao trong $sourcePath"
    Add-Content -Path $logPath -Value $warningMsg -Encoding UTF8
    Write-Warning $warningMsg
    exit
}

foreach ($folder in $folders) {
    $rarFile = Join-Path -Path $sourcePath -ChildPath "$($folder.Name).rar"
    
    try {
        # Nén thư mục thành file RAR (không mật khẩu)
        & $winrarPath a -r -ep1 -m5 -inul "$rarFile" "$($folder.FullName)" | Out-Null
        
        # Kiểm tra xem file RAR đã được tạo thành công chưa
        if (Test-Path $rarFile) {
            # Xóa thư mục gốc nếu nén thành công
            Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Da nen va xoa thu muc: $($folder.Name)"
        } else {
            throw "Khong tim thay file RAR sau khi nen: $rarFile"
        }
    } catch {
        # Ghi lỗi vào file log và hiển thị cảnh báo
        $errorMsg = "$(Get-Date) - Loi khi nen thu muc $($folder.Name): $_"
        Add-Content -Path $logPath -Value $errorMsg -Encoding UTF8
        Write-Warning $errorMsg
    }
}