# Thu nhỏ cửa sổ CMD cho vừa vặn
$wshell = New-Object -ComObject wscript.shell
$wshell.AppActivate('Windows PowerShell') | Out-Null
Start-Sleep -Milliseconds 300

# Chỉnh lại kích thước cửa sổ CMD
$rawUI = $Host.UI.RawUI
$rawUI.WindowSize = New-Object System.Management.Automation.Host.Size(80, 20)   # Số ký tự ngang x dọc
$rawUI.BufferSize = New-Object System.Management.Automation.Host.Size(80, 100)

# Nếu chưa mở bằng CMD, mở lại bằng cửa sổ CMD mới
if (-not $env:KN_SCRIPT_FROM_CMD) {
    $env:KN_SCRIPT_FROM_CMD = "1"

    if ($MyInvocation.MyCommand.Path) {
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process cmd.exe -ArgumentList "/k set KN_SCRIPT_FROM_CMD=1 && powershell -ExecutionPolicy Bypass -File `"$scriptPath`""
        exit
    } else {
        Write-Host "[!] Script đang được chạy từ Internet (irm). Bỏ qua bước mở lại CMD." -ForegroundColor Yellow
    }
}


function Show-Menu {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Gray
    Write-Host "                        KN SCRIPT                          " -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Gray
    Write-Host ""
    Write-Host "[1] Cài Visual Studio Code" -ForegroundColor Green
    Write-Host "[2] Cài Python (chọn phiên bản)" -ForegroundColor Green
    Write-Host "[3] Cài Extension cho VS Code" -ForegroundColor Green
    Write-Host "[4] Active Windows/Office" -ForegroundColor Yellow
    Write-Host "[5] Tải Office Deployment Tool" -ForegroundColor Green
    Write-Host "[6] Tải Config Office" -ForegroundColor Green
    Write-Host "[7] Cài Office (Tự động, không cần [5] và [6])" -ForegroundColor Green
    Write-Host "[8] Cài LocalSend (Gửi tập tin qua mạng LAN)" -ForegroundColor Green
    Write-Host ""
    Write-Host "[0] Thoát bằng tay đi =)))" -ForegroundColor Red
    Write-Host ""
}

function Install-VSCode {
    Write-Host ">> Đang tải và cài đặt VS Code..." -ForegroundColor Cyan
    $url = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user"
    $path = "$env:TEMP\vscode.exe"
    Invoke-WebRequest $url -OutFile $path
    Start-Process $path -ArgumentList "/silent","/mergetasks=!runcode" -Wait
    Write-Host "[✓] Đã cài VS Code!" -ForegroundColor Green
    Pause
}

function Install-Python {
    while ($true) {
        Clear-Host
        Write-Host "============== Chọn phiên bản Python ==============" -ForegroundColor Cyan
        Write-Host "[1] Python 3.12.3"
        Write-Host "[2] Python 3.11.9"
        Write-Host "[3] Python 3.10.14"
        Write-Host "[4] Python 3.9.19"
        Write-Host "[5] Python 3.8.19"
        Write-Host "[6] Python 3.7.17"
        Write-Host "[7] Python 3.6.8"
        Write-Host "[8] Python 3.5.4"
        Write-Host "[9] Python 2.7.18"
        Write-Host "[0] Quay lại menu chính"
        Write-Host ""

        $subChoice = Read-Host "Chọn phiên bản để cài [1-9] hoặc [0] để quay lại"

        $pythonVersions = @{
            "1" = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
            "2" = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
            "3" = "https://www.python.org/ftp/python/3.10.14/python-3.10.14-amd64.exe"
            "4" = "https://www.python.org/ftp/python/3.9.19/python-3.9.19-amd64.exe"
            "5" = "https://www.python.org/ftp/python/3.8.19/python-3.8.19-amd64.exe"
            "6" = "https://www.python.org/ftp/python/3.7.17/python-3.7.17-amd64.exe"
            "7" = "https://www.python.org/ftp/python/3.6.8/python-3.6.8-amd64.exe"
            "8" = "https://www.python.org/ftp/python/3.5.4/python-3.5.4-amd64.exe"
            "9" = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.amd64.msi"
        }

        if ($subChoice -eq "0") { break }

        if ($pythonVersions.ContainsKey($subChoice)) {
            $url = $pythonVersions[$subChoice]
            $file = "$env:TEMP\python-installer.exe"

            Write-Host ">> Đang tải Python từ $url..." -ForegroundColor Cyan
            Invoke-WebRequest $url -OutFile $file

            Write-Host ">> Đang cài đặt Python..." -ForegroundColor Yellow
            if ($url.EndsWith(".msi")) {
                Start-Process msiexec.exe -ArgumentList "/i `"$file`" /qn" -Wait
            } else {
                Start-Process -FilePath $file -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
            }

            Write-Host "[✓] Đã cài Python!" -ForegroundColor Green
            Pause
        } else {
            Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red
            Pause
        }
    }
}

function Install-Extensions {
    $code = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd"
    if (-Not (Test-Path $code)) {
        Write-Host "[!] Không tìm thấy VS Code! Hãy cài đặt trước." -ForegroundColor Red
        Pause
        return
    }

    $extensions = @(
        "ms-python.python",
        "esbenp.prettier-vscode",
        "ms-vscode.cpptools"
    )

    Write-Host ">> Đang cài các extension cho VS Code..." -ForegroundColor Cyan
    foreach ($ext in $extensions) {
        & $code --install-extension $ext --force
    }
    Write-Host "[✓] Đã cài xong extension!" -ForegroundColor Green
    Pause
}

function Activate-Windows {
    Write-Host ">> Đang mở CMD để chạy script kích hoạt..." -ForegroundColor Yellow
    Start-Process cmd.exe -ArgumentList '/k powershell -nop -ep bypass -c "irm https://massgrave.dev/get | iex"'
}

function Download-OfficeDeploymentTool {
    $url = "https://github.com/knhuongnoi1407/office-devployment-tool-download/raw/main/officedeploymenttool_18827-20140.zip"
    $zipPath = "$env:TEMP\officedeploymenttool.zip"
    $extractPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

    Write-Host ">> Đang tải Office Deployment Tool từ GitHub..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Không thể tải Office Deployment Tool." -ForegroundColor Red
        Write-Host "Chi tiết lỗi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Đang giải nén Office Deployment Tool vào thư mục Downloads..." -ForegroundColor Yellow
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    Write-Host "[✓] Đã tải và giải nén Office Deployment Tool tại: $extractPath" -ForegroundColor Green
    Pause
}



function Download-Config {
    while ($true) {
        Clear-Host
        Write-Host "==========      Chọn Phiên Bản Office      ==========" -ForegroundColor Cyan
        Write-Host "[1] Office Professional Plus 2019" -ForegroundColor Green
        Write-Host "[2] Office Professional Plus 2021" -ForegroundColor Green
        Write-Host "[3] Office Professional Plus 2024" -ForegroundColor Green
        Write-Host "[0] Quay lại menu chính" -ForegroundColor Red
        $yearChoice = Read-Host "Chọn [1-3] hoặc [0] để quay lại"
        if ($yearChoice -eq "0") { break }

        $yearMap = @{
            "1" = "2019"
            "2" = "2021"
            "3" = "2024"
        }

        if (-not $yearMap.ContainsKey($yearChoice)) {
            Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red
            Pause
            continue
        }

        $year = $yearMap[$yearChoice]

        Write-Host "`n--- Chọn kiến trúc ---" -ForegroundColor Cyan
        Write-Host "[1] 32bit (x86)" -ForegroundColor Green
        Write-Host "[2] 64bit (x64)" -ForegroundColor Green
        $archChoice = Read-Host "Chọn [1-2]"

        $archMap = @{
            "1" = "32bit"
            "2" = "64bit"
        }

        if (-not $archMap.ContainsKey($archChoice)) {
            Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red
            Pause
            continue
        }

        $arch = $archMap[$archChoice]
        $filename = "Config-office-$year-$arch.xml"
        $url = "https://raw.githubusercontent.com/knhuongnoi1407/config-office-download/main/$filename"
        $dest = [Environment]::GetFolderPath("UserProfile") + "\Downloads\$filename"

        Write-Host ">> Đang tải $filename từ GitHub..." -ForegroundColor Cyan
        Invoke-WebRequest $url -OutFile $dest

        Write-Host "[✓] Đã lưu tại: $dest" -ForegroundColor Green
        Pause
    }
}

function Install-Office {
    Clear-Host
    Write-Host "========== Cài Đặt Office ==========" -ForegroundColor Cyan
    Write-Host "[1] Office Professional Plus 2019" -ForegroundColor Green
    Write-Host "[2] Office Professional Plus 2021" -ForegroundColor Green
    Write-Host "[3] Office Professional Plus 2024" -ForegroundColor Green
    Write-Host "[0] Quay lại" -ForegroundColor Red
    $yearChoice = Read-Host "Chọn [1-3] hoặc [0] để quay lại"
    if ($yearChoice -eq "0") { return }

    $yearMap = @{ "1" = "2019"; "2" = "2021"; "3" = "2024" }

    if (-not $yearMap.ContainsKey($yearChoice)) {
        Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red
        Pause
        return
    }

    $year = $yearMap[$yearChoice]

    Write-Host "`n--- Chọn kiến trúc ---" -ForegroundColor Cyan
    Write-Host "[1] 32bit (x86)" -ForegroundColor Green
    Write-Host "[2] 64bit (x64)" -ForegroundColor Green
    $archChoice = Read-Host "Chọn [1-2]"

    $archMap = @{ "1" = "32bit"; "2" = "64bit" }

    if (-not $archMap.ContainsKey($archChoice)) {
        Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red
        Pause
        return
    }

    $arch = $archMap[$archChoice]
    $folderName = "office$year"

    while ($true) {
        $driveLetter = Read-Host "Nhập ổ đĩa bạn muốn lưu (ví dụ: C, D, E)"
        if (Test-Path "$driveLetter`:") {
            break
        } else {
            Write-Host "[X] Ổ đĩa không tồn tại! Vui lòng nhập lại." -ForegroundColor Red
        }
    }

    $officeDir = "$driveLetter`:\$folderName"
    $configFileName = "Config-office-$year-$arch.xml"
    $configURL = "https://raw.githubusercontent.com/knhuongnoi1407/config-office-download/main/$configFileName"
    $odtZipURL = "https://github.com/knhuongnoi1407/office-devployment-tool-download/raw/main/officedeploymenttool_18827-20140.zip"
    $odtZipPath = "$env:TEMP\officedeploymenttool.zip"

    if (Test-Path $officeDir) {
        Write-Host "[!] Thư mục $officeDir đã tồn tại!" -ForegroundColor Yellow
        Write-Host "[1] Ghi đè toàn bộ (xoá thư mục cũ và tạo mới)" -ForegroundColor Red
        Write-Host "[2] Sử dụng thư mục hiện có" -ForegroundColor Green
        Write-Host "[3] Nhập ổ đĩa khác" -ForegroundColor Cyan
        $overwriteChoice = Read-Host "Chọn hành động [1-3]"
        switch ($overwriteChoice) {
            "1" { Remove-Item -Path $officeDir -Recurse -Force; New-Item -Path $officeDir -ItemType Directory -Force | Out-Null }
            "2" { Write-Host "[i] Sử dụng thư mục hiện có: $officeDir" -ForegroundColor Gray }
            "3" { Install-Office; return }
            Default { Write-Host "[X] Lựa chọn không hợp lệ. Thoát." -ForegroundColor Red; Pause; return }
        }
    } else {
        Write-Host ">> Đang tạo thư mục $folderName tại ổ đĩa $driveLetter..." -ForegroundColor Yellow
        New-Item -Path $officeDir -ItemType Directory -Force | Out-Null
    }

    Write-Host ">> Đang tải config Office ($configFileName)..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest $configURL -OutFile "$officeDir\$configFileName" -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Không thể tải file cấu hình!" -ForegroundColor Red
        Write-Host "Chi tiết lỗi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Đang tải Office Deployment Tool (.zip)..." -ForegroundColor Cyan
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest $odtZipURL -OutFile $odtZipPath -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Không thể tải Office Deployment Tool (.zip)!" -ForegroundColor Red
        Write-Host "Chi tiết lỗi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Đang giải nén Office Deployment Tool vào thư mục $folderName..." -ForegroundColor Yellow
    try {
        Expand-Archive -Path $odtZipPath -DestinationPath $officeDir -Force
        Remove-Item $odtZipPath -Force
    } catch {
        Write-Host "[X] Lỗi khi giải nén Office Deployment Tool!" -ForegroundColor Red
        Pause
        return
    }

    Write-Host "`n[✓] Giải nén xong công cụ Office Deployment Tool." -ForegroundColor Green

    $exePath = Join-Path $officeDir "officedeploymenttool_18827-20140.exe"
    $setupPath = Join-Path $officeDir "setup.exe"

    if (Test-Path $exePath) {
        Write-Host "`n[!] Đang mở file ODT để người dùng chọn thư mục lưu và nhấn Accept..." -ForegroundColor Yellow
        Start-Process -FilePath $exePath

        Write-Host ">> Sau khi cửa sổ hiện ra, hãy chọn thư mục: This PC > $driveLetter > $folderName" -ForegroundColor Cyan
        Write-Host ">> Và nhấn nút Accept để bung file setup.exe." -ForegroundColor Cyan
        Write-Host ">> Đang đợi file setup.exe được bung ra..." -ForegroundColor Cyan

        $timeout = 60
        $elapsed = 0
        while (-not (Test-Path $setupPath) -and $elapsed -lt $timeout) {
            Start-Sleep -Seconds 1
            $elapsed++
        }

        if (Test-Path $setupPath) {
            Write-Host "[✓] setup.exe đã sẵn sàng!" -ForegroundColor Green

            $guidePath = Join-Path $officeDir "Hướng dẫn sử dụng.txt"
            $guideContent = @(
                "Hướng dẫn cài đặt Office thủ công:",
                "",
                "1. Mở Command Prompt (CMD) bằng quyền Administrator",
		"2. Gõ lệnh chuyển ổ đĩa:",
		"   ${driveLetter}",
                "3. Gõ lệnh sau để chuyển thư mục:",
                "   cd ${driveLetter}:\$folderName",
                "4. Gõ lệnh để bắt đầu cài đặt:",
                "   setup.exe /configure $configFileName"
            )
            $guideContent | Set-Content -Path $guidePath -Encoding UTF8

            Write-Host "`n[✓] Đã tạo file 'Hướng dẫn sử dụng.txt' trong thư mục $folderName" -ForegroundColor Green
        } else {
            Write-Host "[X] Không tìm thấy setup.exe sau 60 giây. Vui lòng kiểm tra lại thao tác!" -ForegroundColor Red
        }
    } else {
        Write-Host "[X] Không tìm thấy file officedeploymenttool_18827-20140.exe!" -ForegroundColor Red
    }

    Pause
}

function Install-LocalSend {
    Write-Host ">> Đang tải LocalSend (Windows)... " -ForegroundColor Cyan
    $url = "https://github.com/localsend/localsend/releases/download/v1.17.0/LocalSend-1.17.0-windows-x86-64.exe"
    $file = "$env:TEMP\LocalSend.exe"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $headers = @{ "User-Agent" = "Mozilla/5.0" }
        Invoke-WebRequest -Uri $url -OutFile $file -Headers $headers -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Lỗi khi tải LocalSend: $($_.Exception.Message)" -ForegroundColor Red
        Pause
        return
    }

    Write-Host ">> Đang chạy LocalSend..." -ForegroundColor Yellow
    Start-Process $file
    Write-Host "[✓] Đã mở LocalSend!" -ForegroundColor Green
    Pause
}




function Main {
    while ($true) {
        Show-Menu
        $choice = Read-Host "Chọn một tùy chọn [0-7]"

        switch ($choice) {
            "1" { Install-VSCode }
            "2" { Install-Python }
            "3" { Install-Extensions }
            "4" { Activate-Windows }
            "5" { Download-OfficeDeploymentTool }
            "6" { Download-Config }
            "7" { Install-Office }
            "8" { Install-LocalSend }
            "0" { Write-Host ">> Thoát chương trình..." -ForegroundColor Gray; break }
            Default { Write-Host "[X] Lựa chọn không hợp lệ!" -ForegroundColor Red; Pause }
        }
    }
}

Main
Write-Host "`nNhấn phím bất kỳ để thoát..." -ForegroundColor DarkGray
Pause
