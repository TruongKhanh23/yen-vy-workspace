# convert_excel.ps1

# Lấy ngày hôm nay để làm tên mặc định
$today = Get-Date -Format "dd-MM-yyyy"

# Hỏi người dùng tên folder nguồn
$srcFolderInput = Read-Host "Nhập tên folder nguồn (mặc định: $today)"
if ([string]::IsNullOrWhiteSpace($srcFolderInput)) {
    $srcFolderInput = $today
}

# Thư mục nguồn và đích
$src = Join-Path $PSScriptRoot $srcFolderInput
$dst = Join-Path $PSScriptRoot ("$srcFolderInput-Converted")

# Kiểm tra thư mục nguồn
if (-not (Test-Path $src)) {
    Write-Host "Thư mục nguồn $src không tồn tại!"
    Pause
    exit
}

# Tạo thư mục đích nếu chưa có
if (-not (Test-Path $dst)) {
    New-Item -ItemType Directory -Path $dst | Out-Null
}

# Tạo Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Duyệt tất cả file xlsx trong thư mục nguồn
Get-ChildItem -Path $src -Filter *.xlsx | ForEach-Object {
    Write-Host "Converting $($_.Name)..."
    $wb = $excel.Workbooks.Open($_.FullName)
    $sheet = $wb.Sheets.Item(1)

    # Định dạng các cột theo yêu cầu
    $sheet.Columns.Item(2).NumberFormat = "dd/mm/yyyy"
    $sheet.Columns.Item(7).NumberFormat = '_(* #.##0_);_(* (#.##0);_(* "0"_);_(@_)'
    $sheet.Columns.Item(8).NumberFormat = '_(* #.##0_);_(* (#.##0);_(* "-"??_);_(@_)'
    $sheet.Columns.Item(22).NumberFormat = "@"
    $sheet.Columns.Item(30).NumberFormat = "@"   # đặt kiểu Text
    $sheet.Columns.Item(32).NumberFormat = "@"   # đặt kiểu Text

    $lastRow = $sheet.UsedRange.Rows.Count
    $range = $sheet.Range("AD2:AD$lastRow")      # cột 30 = AD

    foreach ($cell in $range) {
        if ($cell.Value2 -ne $null) {
            $date = [datetime]::FromOADate($cell.Value2)  # nếu ô hiện tại là Excel date
            $cell.Value2 = $date.ToString("dd/MM/yyyy")    # chuyển thành string
        }
    }

    # Gán giá trị 0 cho toàn bộ cột 7
    $usedRange = $sheet.UsedRange
    $lastRow = $usedRange.Rows.Count
    $sheet.Range("G2:G$lastRow").Value2 = 0   # từ dòng 2 đến hết (giữ header dòng 1)

    # ✅ Ép kiểu lại dòng đầu tiên sau khi xử lý xong toàn bộ
    $headerRow = $sheet.Rows.Item(1)
    $headerRow.NumberFormat = "General"

    # Lưu sang xls
    $newName = Join-Path $dst ($_.BaseName + ".xls")
    $wb.SaveAs($newName, 56)
    $wb.Close($false)
}

$excel.Quit()
Write-Host "Done!"
Pause
