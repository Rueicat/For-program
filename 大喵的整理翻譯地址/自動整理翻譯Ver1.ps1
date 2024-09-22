cls
Write-Host @"
##自動批次整理翻譯Ver1    by RueiKato

    -2024.9.21 -依照大喵需求建置自動化程式

##第一次使用請按 YES 授權給程式做系統級別操作

"@
Set-ExecutionPolicy unrestricted -Scope Process

#檢查安裝ImportExcel
if (-not (Get-module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue)){
    Install-Module -Name ImportExcel -Scope CurrentUser
    Write-Host "ImportExcel has installed"
} else {
    Write-Host "ImportExcel has been installed already."
}

#創建GUI讓user選擇需要處理的、excel檔案所在資料夾
Add-Type -AssemblyName System.Windows.Forms

$folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowserDialog.SelectedPath = $PSScriptRoot
$folderBrowserDialog.Description = "請選擇資料夾(批次處理excel所在的資料夾)"
$result = $folderBrowserDialog.ShowDialog()

#確定有選擇資料夾後, 處理files in the select folder
if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $selectFolder = $folderBrowserDialog.SelectedPath

    $OutputPath = Join-Path $PSScriptRoot "result.xlsx"

    #Get all the excel files from select folder
    $ExcelFiles = gci -Path $selectFolder -Filter "*.xlsx"

    #init data shell
    $ProcessedData = @()
    $UnprocessedData = @()

    #處理每個excel文件
    foreach ($File in $ExcelFiles){
        $FileName = $File.FullName

        $ExcelPackage = Open-ExcelPackage -Path $FileName
        $SheetNames = $ExcelPackage.Workbook.Worksheets | Select-Object -ExpandProperty Name

        foreach ($SheetName in $SheetNames){
            $Worksheet = $ExcelPackage.Workbook.Worksheets[$SheetName]

            if ($Worksheet -eq $null){
                continue
            }

            $A3 = $Worksheet.Cells["A3"].Text.Trim()
            $A9 = $Worksheet.Cells["A9"].Text.Trim()

            Write-Host "Processing File: $FileName, Sheet: $SheetName"
            Write-Host "A3: '$A3'"
            Write-Host "A9: '$A9'"

            if ($A3 -match "公司別[:：]" -and $A9 -match "單位名稱"){
                $CompanyName = $A3

                $TotalRow = $null
                for ($row = 10; $row -le $Worksheet.Dimension.End.Row; $row++){
                    $cellValue = $Worksheet.Cells["A$row"].Text
                    if ($cellValue -match ".*總\s+計.*"){
                        $TotalRow = $row
                        break
                    }
                }
                if ($TotalRow -ne $null){
                    $StartRow = 10
                    $EndRow = $TotalRow -1
                } else {
                    #如果沒有"總　　　　計", 則處理到最後一筆資料(這邊覺得會有問題, 先這樣寫)
                    $StartRow = 10
                    $EndRow = $Worksheet.Dimension.End.Row
                }

                #抓數據, and past them to the new sheet[0] in new file
                for ($row = $StartRow; $row -le $EndRow; $row++){
                    $UnitName = $Worksheet.Cells["A$row"].Text
                    $UnitNameEng = $Worksheet.Cells["B$row"].Text
                    $InsuranceLocation = $Worksheet.Cells["C$row"].Text
                    $InsuranceLocationEng = $Worksheet.Cells["D$row"].Text
                    $PropertyAmount = $Worksheet.Cells["E$row"].Text

                    $DataEntry = [PSCustomObject]@{
                        "檔案名稱" = $FileName
                        "公司別" = $CompanyName
                        "營業別" = $SheetName
                        "單位名稱" = $UnitName
                        "單位名稱(英文)" = $UnitNameEng
                        "保險標的處所" = $InsuranceLocation
                        "保險標的處所(英文)" = $InsuranceLocationEng
                        "財產金額" = $PropertyAmount
                    }
                    $ProcessedData += $DataEntry
                }
            } else {
                $UnprocessedEntry = [PSCustomObject]@{
                    "檔案名稱" = $FileName
                    "工作表名稱" = $SheetName
                }
                $UnprocessedData += $UnprocessedEntry
            }

        }
        $ExcelPackage.Dispose()
    }

    #寫入processedData to newsheet[0]
    $ProcessedData | Export-Excel -Path $OutputPath -WorksheetName "已處理資料" -AutoSize -BoldTopRow
    #寫入UnprocessedData to newsheet[1]
    if ($UnprocessedData.Count -gt 0){
        $UnprocessedData | Export-Excel -Path $OutputPath -WorksheetName "未處理資料報告" -AutoSize -BoldTopRow -Append
    } else {
        # 如果没有未處理的data，還是創個sheet
        $UnprocessedHeaders = @{"檔案名稱" = $null; "工作表名稱" = $null}
        [PSCustomObject]$UnprocessedHeaders | Export-Excel -Path $OutputPath -WorksheetName "未處理資料報告" -AutoSize -BoldTopRow -Append
    }
    Write-Host "資料處理完成, 結果已保存到 $OutputPath"

} else {
    Write-Host "你沒有選擇資料夾"
}
Write-Host ""
Write-Host ""
Write-Host "請按任意鍵結束程式"
[void][System.Console]::ReadKey($true)