param(
    [Parameter(Position = 0)]
    [string[]]$Path = ".\*.xls*", # ファイルパス
    [string]$BookName, # ブック名
    [string]$SheetName, # シート名
    [string]$Range, # セル範囲
    [string]$Pattern # 検索文字列
)

if($Pattern -eq "" -and $BookName -eq "" -and $SheetName -eq "" -and $Range -eq ""){
  "検索条件を入力してください"
  $BookName = Read-Host "ブック名"
  $SheetName = Read-Host "シート名"
  $Range = Read-Host "セル範囲"
  $Pattern = Read-Host "検索文字列"
}

# 結果CSV
$result = New-Item ".\result.csv" -Type File -Force

# 対象ブックを取得
$books = Get-Item -Path $Path
$books = $books | ? Name -match $BookName

# ヘッダー
$line = "ブック名,シート名,セル番地,値"
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
        # 対象セル範囲を取得
        if($Range -eq ""){
            $SelectedRange = $sheet.UsedRange
        } else {
            $SelectedRange = $sheet.Range("ZZ99,$Range")
        }

        # 対象セルを取得
        try{
            # 定数
            $constants = $SelectedRange.SpecialCells(
                [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeConstants, ## 定数
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlNumbers + ## 数値
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlTextValues ## テキスト
            )
        } catch{
            $constants = $null
        }
        try{
            # 数式
            $formulas = $SelectedRange.SpecialCells(
                [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeFormulas, ## 数式
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlNumbers + ## 数値
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlTextValues + ## テキスト
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlLogical ## 論理値
            )
        } catch{
            $formulas = $null
        }
        # 対象セル = 定数 + 数式
        $cells = $constants + $formulas

        # 対象セル毎に繰り返し
        foreach($cell in $cells){        
            # セルのテキストを取得
            $text = $cell.Text -Replace "`n","``n"
            if($text -ne "" -and $text -match $Pattern){
                $line = $book.Name + "," + $sheet.Name + "," + [string]$cell.Address($false, $false) + "," + $text
                $line | Write-Output
                $line | Add-Content $result.FullName
            }
        }
    }
    $book.Close(0)
}

# エクセルを終了
$excel.Quit()

# リソース解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
