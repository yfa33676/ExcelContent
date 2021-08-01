function Get-ExcelContent{
  param(
      [Parameter(Position = 0)]
      [string[]]$Path = ".\*.xls*", # ファイルパス
      [string]$Pattern = "" , # 検索文字列
      [string]$BookName = "", # ブック名
      [string]$SheetName = "", # シート名
      [string]$Range = "" # セル範囲
  )

  ## 結果CSV
  #$result = New-Item ".\result.csv" -Type File -Force

  # 対象ブックを取得
  $books = Get-Item -Path $Path
  $books = $books | ? Name -match $BookName


  # エクセルを起動
  $excel = New-Object -ComObject Excel.Application

  # ヘッダー
  $line = "ブック名,シート名,セル番地,値"
  $line | Write-Output
  #$line | Add-Content $result.FullName

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
                  #$line | Add-Content $result.FullName
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
}

function Set-ExcelContent{
  param(
      [Parameter(Position = 0, Mandatory)]
      [string]$ValuePath # 設定値CSVのパス
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
}