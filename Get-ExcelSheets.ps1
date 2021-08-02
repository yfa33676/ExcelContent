param(
    [Parameter(Position = 0)]
    [string[]]$Path = ".\*.xls*", # ファイルパス
    [string]$BookName = "", # ブック名
    [string]$SheetName = "" # シート名
)

if($BookName -eq "" -and $SheetName -eq ""){
  "検索条件を入力してください"
  $BookName = Read-Host "ブック名"
  $SheetName = Read-Host "シート名"
}

# 結果CSV
$result = New-Item ".\result.csv" -Type File -Force

# 対象ブックを取得
$books = Get-Item -Path $Path
$books = $books | ? Name -match $BookName

# ヘッダー
$line = "ブック名,シート名"
$line | Write-Output
$line | Add-Content $result.FullName

# エクセルを起動
$excel = New-Object -ComObject Excel.Application
  

# ブック毎に繰り返し
for($i = 0; $i -lt $books.Count; $i++){
    # ブックを読み取り専用で開く
    $book = $excel.Workbooks.Open($books[$i].FullName, 0, $true)

    # 対象シートを取得
    $sheets = $excel.WorkSheets | ? Name -match $SheetName

    # シート毎に繰り返し
    foreach($sheet in $sheets){
        $line = $book.Name + "," + $sheet.Name
        $line | Write-Output
        $line | Add-Content $result.FullName
    }
    $book.Close(0)
}

# エクセルを終了
$excel.Quit()

# リソース解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
