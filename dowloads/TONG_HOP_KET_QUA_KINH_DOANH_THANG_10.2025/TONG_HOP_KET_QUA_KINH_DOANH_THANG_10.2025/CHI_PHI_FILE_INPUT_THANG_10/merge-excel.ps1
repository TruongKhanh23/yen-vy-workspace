# merge_data.ps1
# PowerShell script tổng hợp dữ liệu Excel từ nhiều file
# Hỏi dòng bắt đầu/kết thúc, sheet, kênh, loại dữ liệu
# Hỏi cột cần lấy (Store, Amount)
# Tự tạo tên file tổng hợp dựa trên lựa chọn người dùng và tháng từ file đầu tiên

# --- Hỏi người dùng dòng bắt đầu và kết thúc ---
$startRow = Read-Host "Nhap dong du lieu bat dau"
$endRow = Read-Host "Nhap dong du lieu ket thuc"

# --- Hỏi tên sheet ---
$sheetName = Read-Host "Nhap ten sheet can xu ly"

# --- Chọn kênh ---
$channels = @("ZALOPAY","SHOPEE","GRAB","XANHSM","VILL","RYO","TIEN_MAT","BE", "VNPAY", "MOMO")
Write-Host "`nBan dang xu ly cho kenh nao?"
for ($i=0; $i -lt $channels.Count; $i++) {
    Write-Host "$($i+1). $($channels[$i])"
}
do {
    $channelIndex = Read-Host "Nhap so tuong ung (1-$($channels.Count))"
} while (-not ($channelIndex -as [int]) -or $channelIndex -lt 1 -or $channelIndex -gt $channels.Count)
$channel = $channels[$channelIndex - 1]
Write-Host "Ban chon kenh: $channel"

# --- Chọn loại dữ liệu ---
$dataTypes = @("DOANH_THU","CHI_PHI")
Write-Host "`nBan dang xu ly LOAI DU LIEU nao?"
for ($i=0; $i -lt $dataTypes.Count; $i++) {
    Write-Host "$($i+1). $($dataTypes[$i])"
}
do {
    $dataTypeIndex = Read-Host "Nhap so tuong ung (1-2)"
} while (-not ($dataTypeIndex -as [int]) -or $dataTypeIndex -lt 1 -or $dataTypeIndex -gt 2)
$dataType = $dataTypes[$dataTypeIndex - 1]
Write-Host "Ban chon loai du lieu: $dataType"

# --- Hỏi cột cần lấy ---
function Get-ColumnNumber($prompt) {
    do {
        $input = Read-Host $prompt
        if ($input -match '^[A-Z]+$') {
            # Chuyển chữ thành số cột (A=1, B=2,...)
            $colNum = 0
            $letters = $input.ToCharArray()
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

$colStore = Get-ColumnNumber "Nhap cot chua Ten cua hang (chu cai hoac so, VD: C hoac 3)"
$colAmount = Get-ColumnNumber "Nhap cot chua So tien (chu cai hoac so, VD: D hoac 4)"
Write-Host "Ban chon cot Store: $colStore, cot Amount: $colAmount"

# --- Lấy folder hiện tại ---
$folder = $PSScriptRoot

# --- Lấy file đầu tiên trong folder ---
$firstFile = Get-ChildItem -Path $folder -Filter "*.xlsx" | Where-Object { $_.BaseName -ne "SHOPEE_Merged" } | Select-Object -First 1
if (-not $firstFile) {
    Write-Host "Khong tim thay file Excel nao trong folder!"
    exit
}

# --- Lấy mm.yyyy từ tên file đầu tiên ---
$dateParts = $firstFile.BaseName -split '\.'  # ["dd","mm","yyyy"]
$monthYear = "$($dateParts[1]).$($dateParts[2])"
$thangPart = "THANG_$monthYear"

# --- Tạo tên file tổng hợp ---
$outputFileName = "${dataType}_${channel}_${thangPart}.xlsx"
$outputFile = Join-Path $folder $outputFileName
Write-Host "`nFile tong hop se duoc tao: $outputFile`n"

# --- Tạo Excel COM object ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# --- Tạo workbook mới để lưu kết quả ---
$wbOut = $excel.Workbooks.Add()
$wsOut = $wbOut.Sheets.Item(1)

# --- Ghi header ---
$wsOut.Cells.Item(1,1) = "Ngay"
$wsOut.Cells.Item(1,2) = "Ten cua hang"
$wsOut.Cells.Item(1,3) = "So tien"

$rowOut = 2

# --- Lặp qua tất cả file Excel trong folder (không phải file tổng hợp) ---
Get-ChildItem -Path $folder -Filter "*.xlsx" | Where-Object { $_.FullName -ne $outputFile } | ForEach-Object {
    $file = $_.FullName
    $fileName = $_.BaseName

    # Chuyển dd.mm.yyyy -> DateTime
    $dateParts = $fileName -split '\.'  # ["dd","mm","yyyy"]
    if ($dateParts.Count -ne 3) { 
        Write-Host "Ten file $fileName khong dung dinh dang dd.mm.yyyy.xlsx, bo qua"
        return
    }
    $day = [int]$dateParts[0]
    $month = [int]$dateParts[1]
    $year = [int]$dateParts[2]
    $dateObj = Get-Date -Year $year -Month $month -Day $day

    try {
        $wb = $excel.Workbooks.Open($file)
        
        # Kiểm tra sheet theo tên người dùng nhập
        if ($wb.Sheets.Item($sheetName)) {
            $ws = $wb.Sheets.Item($sheetName)
        } else {
            Write-Host "Khong tim thay sheet '$sheetName' trong file $fileName"
            $wb.Close($false)
            return
        }

        # Lấy số dòng thực tế có dữ liệu
        $usedRows = $ws.UsedRange.Rows.Count
        Write-Host "Dang xu ly file: $fileName, so dong co du lieu: $usedRows"

        # Duyệt từ startRow -> endRow
        for ($r = [int]$startRow; $r -le [int]$endRow; $r++) {
            if ($r -gt $usedRows) { break }

            $store = $ws.Cells.Item($r, $colStore).Text
            $amount = $ws.Cells.Item($r, $colAmount).Value2

            if ($store -ne "") {
                $wsOut.Cells.Item($rowOut,1) = $dateObj
                $wsOut.Cells.Item($rowOut,1).NumberFormat = "dd/mm/yyyy"
                $wsOut.Cells.Item($rowOut,2) = $store
                $wsOut.Cells.Item($rowOut,3) = $amount
                $rowOut++
            }
        }

        Write-Host "Da xu ly dong $startRow -> $([math]::Min($endRow,$usedRows)) trong file $fileName"
        $wb.Close($false)
    } catch {
        Write-Host "Khong the mo file $fileName"
    }
}

# --- Auto-fit tất cả cột đã sử dụng ---
$wsOut.UsedRange.Columns.AutoFit()

# --- Lưu và đóng file tổng hợp ---
$wbOut.SaveAs($outputFile)
$wbOut.Close()
$excel.Quit()

Write-Host "`nDa tao file tong hop:" $outputFile
Pause
