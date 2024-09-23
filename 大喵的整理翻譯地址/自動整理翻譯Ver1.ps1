Write-Host @"
##自動批次整理翻譯Ver1    by RueiKato
    -2024.9.24 -建置config.json
               -存在編碼問題未解決
    -2024.9.22 -依需求做修正
                    -顯示檔名路徑調整
                    -除錯(位置對應錯誤的問題)
                    -建置翻譯功能
    -2024.9.21 -依照大喵需求建置自動化程式

##第一次使用請按 YES 授權給程式做系統級別操作

"@
Set-ExecutionPolicy unrestricted -Scope Process

#檢查安裝ImportExcel
if (-not (Get-module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue)){
    Install-Module -Name ImportExcel -Scope CurrentUser
    Import-Module -Name ImportExcel
    Write-Host "ImportExcel has been installed and loaded."
} else {
    Write-Host "ImportExcel has been installed already."
}

#創建GUI讓user選擇需要處理的excel檔案所在資料夾
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
    $ExcelFiles = Get-ChildItem -Path $selectFolder -Filter "*.xlsx"

    #init data shell
    $ProcessedData = @()
    $UnprocessedData = @()

    #設置config 文件取得api and location value (自行設置自己的value)
    #翻譯是使用MS translator API, 可以在Azure上註冊帳號後選擇免費方案有限度的使用
    $configFile = ".\config.json"

    $config = Get-Content $configFile | ConvertFrom-Json

    $YOUR_API_KEY = $config.YOUR_API_KEY
    $YOUR_LOCATION = $config.YOUR_LOCATION

    function translatems($Text, $FromLang, $ToLang){
        $subscriptionKey = $YOUR_API_KEY  
        $endpoint = 'https://api.cognitive.microsofttranslator.com/'
        $location = $YOUR_LOCATION

        $path = '/translate?api-version=3.0'
        $params = "&from=$FromLang&to=$ToLang"

        $uri = $endpoint.TrimEnd('/') + $path + $params

        $body = @(
            @{
                'Text' = $Text
            }
        )

        $headers = @{
            'Ocp-Apim-Subscription-Key' = $subscriptionKey
            'Ocp-Apim-Subscription-Region' = $location
            'Content-Type' = 'application/json'
        }

        function ConvertTo-JsonForceArray($InputObject, $Depth) {
            $json = $InputObject | ConvertTo-Json -Depth $Depth
            if ($InputObject.Count -eq 1 -and $json.StartsWith('{')) {
                return "[$json]"
            } else {
                return $json
            }
        }



        try {
            $bodyJson = ConvertTo-JsonForceArray -InputObject $body -Depth 2
            # Write-Host "翻譯請求的 URI: $uri"
            Write-Host "翻譯文字: $bodyJson"
            
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -ContentType 'application/json' -Body $bodyJson

            if ($response -and $response[0].translations) {
                return $response[0].translations[0].text
            } else {
                Write-Host "翻譯回應為空或格式錯誤"
                return ""
            }
        }
        catch {
            Write-Host "翻譯 API 呼叫失敗: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $errorResponse = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($errorResponse)
                $errorContent = $reader.ReadToEnd()
                Write-Host "錯誤內容: $errorContent"
            }
            return ""
        }
    }

    #處理每個excel文件
    foreach ($File in $ExcelFiles){
        $FileName = Split-Path -Path $File.Fullname -Leaf

        $ExcelPackage = Open-ExcelPackage -Path $File.Fullname
        $SheetNames = $ExcelPackage.Workbook.Worksheets | Select-Object -ExpandProperty Name

        foreach ($SheetName in $SheetNames){
            $Worksheet = $ExcelPackage.Workbook.Worksheets[$SheetName]

            if ($null -eq $Worksheet){
                continue
            }

            $A3 = $Worksheet.Cells["A3"].Text.Trim()
            $A9 = $Worksheet.Cells["A9"].Text.Trim()

            Write-Host ""
            Write-Host "處理的檔案名稱:$FileName`n處理的頁籤: $SheetName"
            Write-Host ""

            if ($A3 -match "公司別[:：]" -and $A9 -match "單位名稱"){
                $CompanyName = $A3

                #最後一筆有效資料的row
                $EndRow = $null
               
                for ($row = 10; $row -le $Worksheet.Dimension.End.Row; $row++){
                    $cellValue = $Worksheet.Cells["A$row"].Text
                
                    if ([string]::IsNullOrWhiteSpace($cellValue)){
                        $EndRow = $row - 1
                        break
                    } else {
                        $EndRow = $row
                    }
                }

                if ($null -eq $EndRow){
                    $EndRow = $Worksheet.Dimension.End.Row
                }

                #檢查遇到"總    計"的問題
                $CellValuePreviousRow = $Worksheet.Cells["A$EndRow"].Text.Trim()
                if ($CellValuePreviousRow -match '^總\s+計$'){
                    $EndRow = $EndRow -1
                    if ($EndRow -lt 10){
                        $EndRow = 10
                    }
                }

                #抓數據, and paste them to the new sheet[0] in new file
                $StartRow = 10     #從row10開始處理
                for ($row = $StartRow; $row -le $EndRow; $row++){
                    $UnitName = $Worksheet.Cells["A$row"].Text
                    $UnitNameEng = translatems -Text $UnitName -FromLang 'zh-Hant' -ToLang 'en'
                    $InsuranceLocation = $Worksheet.Cells["B$row"].Text
                    $InsuranceLocationEng = translatems -Text $InsuranceLocation -FromLang 'zh-Hant' -ToLang 'en'
                    $PropertyAmount = $Worksheet.Cells["C$row"].Text

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
