$wshell = New-Object -ComObject wscript.shell
$wshell.AppActivate('Windows PowerShell') | Out-Null
Start-Sleep -Milliseconds 300

$rawUI = $Host.UI.RawUI
$rawUI.WindowSize = New-Object System.Management.Automation.Host.Size(80, 20)   
$rawUI.BufferSize = New-Object System.Management.Automation.Host.Size(80, 100)

if (-not $env:KN_SCRIPT_FROM_CMD) {
    $env:KN_SCRIPT_FROM_CMD = "1"

    if ($MyInvocation.MyCommand.Path) {
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process cmd.exe -ArgumentList "/k set KN_SCRIPT_FROM_CMD=1 && powershell -ExecutionPolicy Bypass -File `"$scriptPath`""
        exit
    } else {
        Write-Host "[!] Script dang duoc chay tu Internet (irm). Bo qua buoc mo lai CMD." -ForegroundColor Yellow
    }
}


function Show-Menu {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Gray
    Write-Host "                        KN SCRIPT                          " -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Gray
    Write-Host ""
    Write-Host "[1] Cai Visual Studio Code" -ForegroundColor Green
    Write-Host "[2] Cai Python (chon phien ban)" -ForegroundColor Green
    Write-Host "[3] Cai Extension cho VS Code" -ForegroundColor Green
    Write-Host "[4] Active Windows/Office" -ForegroundColor Yellow
    Write-Host "[5] Tai Office Deployment Tool" -ForegroundColor Green
    Write-Host "[6] Tai Config Office" -ForegroundColor Green
    Write-Host "[7] Cai Office (Tu dong, khong can [5] va [6])" -ForegroundColor Yellow
    Write-Host "[8] Cai LocalSend (Gui tap tin qua mang LAN)" -ForegroundColor Green
    Write-Host "[9] Cai WinRAR" -ForegroundColor Green
    Write-Host ""
    Write-Host "[0] Thoat | Exit" -ForegroundColor Red
    Write-Host ""
}

function Install-VSCode {
    Write-Host ">> Dang tai va cai dat VS Code..." -ForegroundColor Cyan
    $url = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user"
    $path = "$env:TEMP\vscode.exe"
    Invoke-WebRequest $url -OutFile $path
    Start-Process $path -ArgumentList "/silent","/mergetasks=!runcode" -Wait
    Write-Host "[✓] Da cai VS Code!" -ForegroundColor Green
    Pause
}

function Install-Python {
    while ($true) {
        Clear-Host
        Write-Host "============== Chon phien ban Python ==============" -ForegroundColor Cyan
        Write-Host "[1] Python 3.12.3"
        Write-Host "[2] Python 3.11.9"
        Write-Host "[3] Python 3.10.14"
        Write-Host "[4] Python 3.9.19"
        Write-Host "[5] Python 3.8.19"
        Write-Host "[6] Python 3.7.17"
        Write-Host "[7] Python 3.6.8"
        Write-Host "[8] Python 3.5.4"
        Write-Host "[9] Python 2.7.18"
        Write-Host "[0] Quay lai menu chinh"
        Write-Host ""

        $subChoice = Read-Host "Chon phien ban de cai [1-9] hoac [0] de quay lai"

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

            Write-Host ">> Dang tai Python từ $url..." -ForegroundColor Cyan
            Invoke-WebRequest $url -OutFile $file

            Write-Host ">> Dang cai dat Python..." -ForegroundColor Yellow
            if ($url.EndsWith(".msi")) {
                Start-Process msiexec.exe -ArgumentList "/i `"$file`" /qn" -Wait
            } else {
                Start-Process -FilePath $file -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
            }

            Write-Host "[✓] Da cai Python!" -ForegroundColor Green
            Pause
        } else {
            Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red
            Pause
        }
    }
}

function Install-Extensions {
    $code = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\bin\code.cmd"
    if (-Not (Test-Path $code)) {
        Write-Host "[!] Khong tim thay VS Code! Hay cai dat truoc." -ForegroundColor Red
        Pause
        return
    }

    $extensions = @(
        "batisteo.vscode-django",
		"battlebas.kivy-vscode",
  		"chrisru.vscode-nightsky",
		"donjayamanne.python-environment-manager",
  		"donjayamanne.python-extension-pack",
        "esbenp.prettier-vscode",
        "github.github-vscode-theme",
        "kevinrose.vsc-python-indent",
        "ms-python.debugpy",
        "ms-python.python",
        "ms-python.vscode-pylance",
        "negokaz.live-server-preview",
        "njpwerner.autodocstring",
        "pkief.material-icon-theme",
        "ritwickdey.liveserver",
        "streetsidesoftware.code-spell-checker",
        "streetsidesoftware.code-spell-checker-vietnamese",
        "visualstudioexptteam.intellicode-api-usage-examples",
        "visualstudioexptteam.vscodeintellicode",
        "wholroyd.jinja"
    )

    Write-Host ">> Dang cai cac extention VS Code..." -ForegroundColor Cyan
    foreach ($ext in $extensions) {
        & $code --install-extension $ext --force
    }
    Write-Host "[✓] Da cai xong extension!" -ForegroundColor Green
    Pause
}

function Activate-Windows {
    Write-Host ">> Dang mo CMD de chay script kich hoat..." -ForegroundColor Yellow
    Start-Process cmd.exe -ArgumentList '/k powershell -nop -ep bypass -c "irm https://get.activated.win | iex"'
}

function Download-OfficeDeploymentTool {
    $url = "https://github.com/knhuongnoi1407/office-devployment-tool-download/raw/main/officedeploymenttool_18827-20140.zip"
    $zipPath = "$env:TEMP\officedeploymenttool.zip"
    $extractPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

    Write-Host ">> Dang tai Office Deployment Tool tu GitHub..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Khong the tai Office Deployment Tool." -ForegroundColor Red
        Write-Host "Chi tiet loi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Dang giai nen Office Deployment Tool vao thu muc Downloads..." -ForegroundColor Yellow
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    Write-Host "[✓] Da tai va giai nen Office Deployment Tool tai: $extractPath" -ForegroundColor Green
    Pause
}



function Download-Config {
    while ($true) {
        Clear-Host
        Write-Host "==========      Chon phien ban Office      ==========" -ForegroundColor Cyan
        Write-Host "[1] Office Professional Plus 2019" -ForegroundColor Green
        Write-Host "[2] Office Professional Plus 2021" -ForegroundColor Green
        Write-Host "[3] Office Professional Plus 2024" -ForegroundColor Green
        Write-Host "[0] Quay lai menu chinh" -ForegroundColor Red
        $yearChoice = Read-Host "Chon [1-3] hoac [0] de quay lai"
        if ($yearChoice -eq "0") { break }

        $yearMap = @{
            "1" = "2019"
            "2" = "2021"
            "3" = "2024"
        }

        if (-not $yearMap.ContainsKey($yearChoice)) {
            Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red
            Pause
            continue
        }

        $year = $yearMap[$yearChoice]

        Write-Host "`n--- Chon kien truc ---" -ForegroundColor Cyan
        Write-Host "[1] 32bit (x86)" -ForegroundColor Green
        Write-Host "[2] 64bit (x64)" -ForegroundColor Green
        $archChoice = Read-Host "Chọn [1-2]"

        $archMap = @{
            "1" = "32bit"
            "2" = "64bit"
        }

        if (-not $archMap.ContainsKey($archChoice)) {
            Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red
            Pause
            continue
        }

        $arch = $archMap[$archChoice]
        $filename = "Config-office-$year-$arch.xml"
        $url = "https://raw.githubusercontent.com/knhuongnoi1407/config-office-download/main/$filename"
        $dest = [Environment]::GetFolderPath("UserProfile") + "\Downloads\$filename"

        Write-Host ">> Dang tai $filename tu GitHub..." -ForegroundColor Cyan
        Invoke-WebRequest $url -OutFile $dest

        Write-Host "[✓] Da luu tai: $dest" -ForegroundColor Green
        Pause
    }
}

function Install-Office {
    Clear-Host
    Write-Host "========== Cai dat Office ==========" -ForegroundColor Cyan
    Write-Host "[1] Office Professional Plus 2019" -ForegroundColor Green
    Write-Host "[2] Office Professional Plus 2021" -ForegroundColor Green
    Write-Host "[3] Office Professional Plus 2024 (khong khuyen dung)" -ForegroundColor Green
    Write-Host "[0] Quay lai" -ForegroundColor Red
    $yearChoice = Read-Host "Chon [1-3] hoac [0] de quay lai"
    if ($yearChoice -eq "0") { return }

    $yearMap = @{ "1" = "2019"; "2" = "2021"; "3" = "2024" }

    if (-not $yearMap.ContainsKey($yearChoice)) {
        Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red
        Pause
        return
    }

    $year = $yearMap[$yearChoice]

    Write-Host "`n--- Chon kien truc ---" -ForegroundColor Cyan
    Write-Host "[1] 32bit (x86)" -ForegroundColor Green
    Write-Host "[2] 64bit (x64)" -ForegroundColor Green
    $archChoice = Read-Host "Chon [1-2]"

    $archMap = @{ "1" = "32bit"; "2" = "64bit" }

    if (-not $archMap.ContainsKey($archChoice)) {
        Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red
        Pause
        return
    }

    $arch = $archMap[$archChoice]
    $folderName = "office$year"

    while ($true) {
        $driveLetter = Read-Host "Nhap o dia ban muon luu (vi du: C, D, E)"
        $driveLetter = $driveLetter.ToUpper()
        if (Test-Path "$driveLetter`:") {
            break
        } else {
            Write-Host "[X] O dia khong ton tai! Vui long nhap lai." -ForegroundColor Red
        }
    }

    $officeDir = "$driveLetter`:\$folderName"
    $configFileName = "Config-office-$year-$arch.xml"
    $configURL = "https://raw.githubusercontent.com/knhuongnoi1407/config-office-download/main/$configFileName"
    $odtZipURL = "https://github.com/knhuongnoi1407/office-devployment-tool-download/raw/main/officedeploymenttool_18827-20140.zip"
    $odtZipPath = "$env:TEMP\officedeploymenttool.zip"

    if (Test-Path $officeDir) {
        Write-Host "[!] Thu muc $officeDir da ton tai!" -ForegroundColor Yellow
        Write-Host "[1] Ghi de toan bo (xoa thu muc cu va tao moi)" -ForegroundColor Red
        Write-Host "[2] Su dung thu muc hien co" -ForegroundColor Green
        Write-Host "[3] Nhap o dia khac" -ForegroundColor Cyan
        $overwriteChoice = Read-Host "Chon hanh dong [1-3]"
        switch ($overwriteChoice) {
            "1" { Remove-Item -Path $officeDir -Recurse -Force; New-Item -Path $officeDir -ItemType Directory -Force | Out-Null }
            "2" { Write-Host "[i] Su dung thu muc hien co: $officeDir" -ForegroundColor Gray }
            "3" { Install-Office; return }
            Default { Write-Host "[X] Lua chon khong hop le. Thoat." -ForegroundColor Red; Pause; return }
        }
    } else {
        Write-Host ">> Dang tao thu muc $folderName tai o dia $driveLetter..." -ForegroundColor Yellow
        New-Item -Path $officeDir -ItemType Directory -Force | Out-Null
    }

    Write-Host ">> Dang tai config Office ($configFileName)..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest $configURL -OutFile "$officeDir\$configFileName" -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Khong the tao file cau hinh!" -ForegroundColor Red
        Write-Host "Chi tiet loi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Dang tai Office Deployment Tool (.zip)..." -ForegroundColor Cyan
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest $odtZipURL -OutFile $odtZipPath -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Khong the tai Office Deployment Tool (.zip)!" -ForegroundColor Red
        Write-Host "Chi tiet loi: $($_.Exception.Message)" -ForegroundColor DarkGray
        Pause
        return
    }

    Write-Host ">> Dang giai ne Office Deployment Tool vao thu muc $folderName..." -ForegroundColor Yellow
    try {
        Expand-Archive -Path $odtZipPath -DestinationPath $officeDir -Force
        Remove-Item $odtZipPath -Force
    } catch {
        Write-Host "[X] Loi khi giai nen Office Deployment Tool!" -ForegroundColor Red
        Pause
        return
    }

    Write-Host "`n[✓] Giai nen cong cu Office Deployment Tool." -ForegroundColor Green

    $exePath = Join-Path $officeDir "officedeploymenttool_18827-20140.exe"
    $setupPath = Join-Path $officeDir "setup.exe"

    if (Test-Path $exePath) {
    Write-Host "`n[!] Dang mo file ODT de nguoi dung chon thu muc va an Accept..." -ForegroundColor Yellow
    Start-Process -FilePath $exePath

    Write-Host ">> Sau khi cua so hien ra, hay nhan: This PC > $driveLetter > $folderName" -ForegroundColor Cyan
    Write-Host ">> Va nhan nut Accept de bung file setup.exe." -ForegroundColor Cyan
    Write-Host ">> Dang doi file setup.exe duoc bung ra..." -ForegroundColor Cyan

    $timeout = 60
    $elapsed = 0
    while (-not (Test-Path $setupPath) -and $elapsed -lt $timeout) {
        Start-Sleep -Seconds 1
        $elapsed++
    }

    if (Test-Path $setupPath) {
        Write-Host "[✓] setup.exe da san sang!" -ForegroundColor Green

        $guidePath = Join-Path $officeDir "Huong dan su dung.txt"
        $guideContent = @(
            "Huong dan cai dat Office thu cong:",
            "",
            "1. Mo Command Prompt (CMD) bang quyen Administrator",
            "2. Go lenh sau de chuyen thu muc:",
            "cd /${driveLetter} ${driveLetter}:\$folderName",
            "3. Go lenh de bat dau cai dat:",
            "setup.exe /configure $configFileName",
			"",
			"",
			"",
			"Ban quyen Script thuoc KN"
        )
        $guideContent | Set-Content -Path $guidePath -Encoding UTF8
        Write-Host "`n[✓] Da tao file 'Huong dan su dung.txt' trong thu muc $folderName" -ForegroundColor Green
    } else {
        Write-Host "[X] Khong tim thay setup.exe sau 60 giay. Vui long kiem tra lai thao tac!" -ForegroundColor Red
    }
    }


    Pause
}

function Install-LocalSend {
    Write-Host ">> Đang tai LocalSend (Windows)... " -ForegroundColor Cyan
    $url = "https://github.com/localsend/localsend/releases/download/v1.17.0/LocalSend-1.17.0-windows-x86-64.exe"
    $file = "$env:TEMP\LocalSend.exe"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $headers = @{ "User-Agent" = "Mozilla/5.0" }
        Invoke-WebRequest -Uri $url -OutFile $file -Headers $headers -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Loi khi tai LocalSend: $($_.Exception.Message)" -ForegroundColor Red
        Pause
        return
    }

    Write-Host ">> Dang chay LocalSend..." -ForegroundColor Yellow
    Start-Process $file
    Write-Host "[✓] Da mo LocalSend!" -ForegroundColor Green
    Pause
}

function Install-WinRAR {
    Clear-Host
    Write-Host "========== Chon phien ban WinRAR ==========" -ForegroundColor Cyan
    Write-Host "[1] 64-bit" -ForegroundColor Green
    Write-Host "[2] 32-bit" -ForegroundColor Green
    Write-Host "[0] Quay lai" -ForegroundColor Red
    $archChoice = Read-Host "Chon [1-2] hoac [0] de quay lai"

    switch ($archChoice) {
        "0" { return }
        "1" { $url = "https://www.rarlab.com/rar/winrar-x64-602.exe" }   # 64-bit
        "2" { $url = "https://www.rarlab.com/rar/winrar-x86-602.exe" }    # 32-bit
        Default { Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red; Pause; return }
    }

    $file = "$env:TEMP\winrar.exe"

    Write-Host ">> Dang tai WinRAR tu $url ..." -ForegroundColor Cyan
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $file -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[X] Loi khi tai WinRAR: $($_.Exception.Message)" -ForegroundColor Red
        Pause
        return
    }

    Write-Host ">> Dang cai WinRAR va tich hop vao Windows..." -ForegroundColor Yellow
    # /S = silent, /Reg = thiết lập mặc định cho ZIP và các định dạng, /D="..." = folder cài
    Start-Process -FilePath $file -ArgumentList "/S /Reg" -Wait

    Write-Host "[✓] Da cai WinRAR va dat lam mac dinh cho ZIP!" -ForegroundColor Green
    Pause
}



function Main {
    while ($true) {
        Show-Menu
        $choice = Read-Host "Chon mot tuy chon [0-9]"

        switch ($choice) {
            "1" { Install-VSCode }
            "2" { Install-Python }
            "3" { Install-Extensions }
            "4" { Activate-Windows }
            "5" { Download-OfficeDeploymentTool }
            "6" { Download-Config }
            "7" { Install-Office }
            "8" { Install-LocalSend }
	    "9" { Install-WinRAR }
            "0" {
				Clear-Host
    			Write-Host "Dang thoat | Exting..." -ForegroundColor Gray
    			Start-Sleep -Milliseconds 500
    			exit
			}
            Default { Write-Host "[X] Lua chon khong hop le!" -ForegroundColor Red; Pause }
        }
    }
}

Main
Write-Host "`nNhan phim bat ki de thoat..." -ForegroundColor DarkGray
Pause







