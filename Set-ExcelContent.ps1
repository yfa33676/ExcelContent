param(
    [Parameter(Position = 0)]
    [string]$ValuePath = ".\result.csv"
)
# 対象ブックを取得
$books = Get-Content -Path $ValuePath | ConvertFrom-Csv | Select-Object ブック名 -Unique | % ブック名 | Get-Item

# エクセルを起動
$excel = New-Object -ComObject Excel.Application

# ブック毎に繰り返し
for($i = 0; $i -lt $books.Count; $i++){
    # ブックを開く
    $book = $excel.Workbooks.Open($books[$i].FullName, 0, $false)
    
    # 対象セルを取得
    $cells = Get-Content -Path $ValuePath | ConvertFrom-Csv | Where-Object ブック名 -eq $book.Name

    # セル毎に繰り返し
    foreach($cell in $cells){
        $cell.ブック名 + "," + $cell.シート名 + "," + $cell.セル番地 + "," + $cell.値
        # 対象シートを選択
        $sheet = $excel.WorkSheets.Item($cell.シート名)
        # 対象セルに値を設定
        $sheet.Cells.Range($cell.セル番地).Value2 = $cell.値 -Replace "``n","`n"
    }

    # ブックを保存して閉じる
    $book.Close(1)
}
# エクセルを終了
$excel.Quit()

# リソース解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null