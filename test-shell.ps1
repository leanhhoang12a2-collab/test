# Đường dẫn tới tệp Excel chứa danh sách các địa chỉ Telnet và cổng
$excelFilePath = "C:\Users\HOANGLA\Downloads\test-shell.xlsx"

# Đọc dữ liệu từ tệp Excel
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item(1)
$range = $worksheet.UsedRange
$rowCount = $range.Rows.Count

# Lặp qua từng hàng trong tệp Excel
for ($i = 1; $i -le $rowCount; $i++) {
    # Lấy giá trị từ cột A (đường dẫn Telnet) và cột B (cổng)
    $telnetAddress = $worksheet.Cells.Item($i, 1).Value2
    $port = $worksheet.Cells.Item($i, 2).Value2

    # Kiểm tra kết nối Telnet
    $telnetResult = Test-NetConnection -ComputerName $telnetAddress -Port $port

    # Kiểm tra kết quả và hiển thị thông báo tương ứng
    if ($telnetResult.TcpTestSucceeded) {
        Write-Host "Kết nối Telnet thành công tới ${telnetAddress}:${port}"
    }
    else {
        Write-Host "Không thể kết nối Telnet tới ${telnetAddress}:${port}"
    }
}

# Đóng tệp Excel và thoát
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
