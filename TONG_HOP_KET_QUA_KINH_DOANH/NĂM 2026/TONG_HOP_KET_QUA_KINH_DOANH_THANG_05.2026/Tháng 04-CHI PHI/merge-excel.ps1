# merge_data.ps1
# Phiên bản hoàn chỉnh (fix chọn kênh, fix đọc nhiều cột)
# Tác giả: ChatGPT (GPT-5)
# Ngày cập nhật: 01/11/2025

# --- Hỏi dòng bắt đầu và kết thúc ---
$startRow = Read-Host "Nhap dong du lieu bat dau"
$endRow = Read-Host "Nhap dong du lieu ket thuc"

# --- Hỏi tên sheet ---
$sheetName = Read-Host "Nhap ten sheet can xu ly"

# --- Chọn kênh ---
$channels = @("ZALOPAY","SHOPEE","GRAB","XANHSM","VILL","RYO","TIEN_MAT","BE","VNPAY","MOMO")
Write-Host "`nBan dang xu ly cho kenh nao?"
for ($i=0; $i -lt $channels.Count; $i++) {
    Write-Host "$($i+1). $($channels[$i])"
}
do {
    $channelIndexStr = Read-Host "Nhap so tuong ung (1-$($channels.Count))"
    if ($channelIndexStr -match '^\d+$') {
        $channelIndex = [int]$channelIndexStr
    } else {
        $channelIndex = 0
    }
} while ($channelIndex -lt 1 -or $channelIndex -gt $channels.Count)
$channel = $channels[$channelIndex - 1]
Write-Host "Ban chon kenh: $channel"

# --- Chọn loại dữ liệu ---
$dataTypes = @("DOANH_THU","CHI_PHI")
Write-Host "`nBan dang xu ly LOAI DU LIEU nao?"
for ($i=0; $i -lt $dataTypes.Count; $i++) {
    Write-Host "$($i+1). $($dataTypes[$i])"
}
do {
    $dataTypeStr = Read-Host "Nhap so tuong ung (1-2)"
    if ($dataTypeStr -match '^\d+$') {
        $dataTypeIndex = [int]$dataTypeStr
    } else {
        $dataTypeIndex = 0
    }
} while ($dataTypeIndex -lt 1 -or $dataTypeIndex -gt $dataTypes.Count)
$dataType = $dataTypes[$dataTypeIndex - 1]
Write-Host "Ban chon loai du lieu: $dataType"

# --- Hàm chuyển chữ cái sang số cột ---
function Get-ColumnNumber($prompt) {
    do {
        $input = Read-Host $prompt
        if ($input -match '^[A-Z]+$') {
            $colNum = 0
            $letters = $input.ToUpper().ToCharArray()
            foreach ($ch in $letters) {
                $colNum = $colNum * 26 + ([int][char]$ch - [int][char]'A' + 1)
            }
        } elseif ($input -match '^\d+$') {
            $colNum = [int]$input
        } else {
            $colNum = $null
        }
    } while (-not $colNum)
    return $colNum
}

# --- Hỏi số lượng cột ---
do {
    $colCount = Read-Host "Nhap so luong cot muon lay (VD: 2, 3, 5...)"
} while (-not ($colCount -as [int]) -or $colCount -lt 1)
$colCount = [int]$colCount

# --- Hỏi từng cột ---
$columns = @()
for ($i=1; $i -le $colCount; $i++) {
    $colLetter = Get-ColumnNumber "Nhap cot thu $i (VD: C hoac 3)"
    $colName = Read-Host "Nhap ten cho cot thu $i (VD: Ten cua hang, So tien, Ma don...)"
    $columns += [PSCustomObject]@{
        Name = $colName
        Index = $colLetter
    }
}

Write-Host "`nBan da chon cac cot:"
$columns | ForEach-Object { Write-Host " - $($_.Name): Cot $($_.Index)" }

# --- Folder hiện tại ---
$folder = $PSScriptRoot

# --- Lấy file đầu tiên ---
$firstFile = Get-ChildItem -Path $folder -Filter "*.xlsx" | Where-Object { $_.BaseName -notmatch "_Merged$" } | Select-Object -First 1
if (-not $firstFile) {
    Write-Host "Khong tim thay file Excel nao trong folder!"
    exit
}

# --- Xác định tháng ---
$dateParts = $firstFile.BaseName -split '\.'
if ($dateParts.Count -ge 3) {
    $monthYear = "$($dateParts[1]).$($dateParts[2])"
} else {
    $monthYear = (Get-Date -Format "MM.yyyy")
}
$thangPart = "THANG_$monthYear"

# --- Tên file tổng hợp ---
$outputFileName = "${dataType}_${channel}_${thangPart}.xlsx"
$outputFile = Join-Path $folder $outputFileName
Write-Host "`nFile tong hop se duoc tao: $outputFile`n"

# --- Tạo Excel COM object ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# --- Workbook output ---
$wbOut = $excel.Workbooks.Add()
$wsOut = $wbOut.Sheets.Item(1)

# --- Header ---
$wsOut.Cells.Item(1,1) = "Ngay"
for ($i=0; $i -lt $columns.Count; $i++) {
    $wsOut.Cells.Item(1, $i + 2) = $columns[$i].Name
}
$rowOut = 2

# --- Duyệt tất cả file ---
Get-ChildItem -Path $folder -Filter "*.xlsx" | Where-Object { $_.FullName -ne $outputFile } | ForEach-Object {
    $file = $_.FullName
    $fileName = $_.BaseName

    $dateParts = $fileName -split '\.'
    if ($dateParts.Count -ne 3) {
        Write-Host "Bo qua $fileName (ten khong dung dd.mm.yyyy)"
        return
    }
    $day = [int]$dateParts[0]
    $month = [int]$dateParts[1]
    $year = [int]$dateParts[2]
    $dateObj = Get-Date -Year $year -Month $month -Day $day

    try {
        $wb = $excel.Workbooks.Open($file)
        $ws = $wb.Sheets.Item($sheetName)
        $null = $ws.UsedRange.Value2   # kích hoạt UsedRange
        $usedRows = $ws.UsedRange.Rows.Count
        Write-Host "Dang xu ly: $fileName ($usedRows dong du lieu)"

        for ($r = [int]$startRow; $r -le [int]$endRow; $r++) {
            if ($r -gt $usedRows) { break }

            $hasData = $false
            $rowValues = @()
            foreach ($col in $columns) {
                $cell = $ws.Cells.Item($r, $col.Index)
                $val = $cell.Value2
                if ($null -eq $val -or $val -eq "") {
                    $val = $cell.Text
                }
                $rowValues += $val
                if ($val -ne "") { $hasData = $true }
            }

            if ($hasData) {
                $wsOut.Cells.Item($rowOut,1) = $dateObj
                $wsOut.Cells.Item($rowOut,1).NumberFormat = "dd/mm/yyyy"
                for ($i=0; $i -lt $rowValues.Count; $i++) {
                    $wsOut.Cells.Item($rowOut, $i + 2) = $rowValues[$i]
                }
                $rowOut++
            }
        }

        Write-Host "→ Da xu ly xong $fileName"
        $wb.Close($false)
    } catch {
        Write-Host "⚠️ Loi khi mo file $fileName"
    }
}

$wsOut.UsedRange.Columns.AutoFit()
$wbOut.SaveAs($outputFile)
$wbOut.Close()
$excel.Quit()

Write-Host "`n✅ Da tao file tong hop:" $outputFile
Pause
