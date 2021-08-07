function Get-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [psobject]$InputObject, # 入力オブジェクト
        [Parameter(ValueFromPipeline, Position = 0)]
        [string]$Path, # ファイルパス
        [string]$Pattern, # 検索文字列
        [string]$BookName, # ブック名
        [string]$SheetName, # シート名
        [string]$Range # セル範囲
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # 対象ブックの取得
        if($Path -ne ""){
            $BookFiles = Get-Item -Path $Path | ? Name -match $BookName
        } else {
            $BookFiles = $InputObject | ? Name -match $BookName
        }

        # ブック毎に繰り返し
        foreach($BookFile in $BookFiles){
            # ブックを読み取り専用で開く
            $book = $excel.Workbooks.Open($BookFile, 0, $true)

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
                        [PSCustomObject]@{ブック名 = $book.Name; シート名 = $sheet.Name; セル番地 = $cell.Address($false, $false); 値 = $text} | Write-Output
                    }
                }
            }
            
            # ブックを閉じる
            $book.Close(0)
        }

        # リソース解放
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }	
    end {
        # エクセルを終了
        $excel.Quit()

        # リソース解放
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Get-ExcelSheets{
    param(
        [Parameter(ValueFromPipeline)]
        [psobject]$InputObject,
        [Parameter(ValueFromPipeline, Position = 0)]
        [string]$Path
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # 対象ブックの取得
        if($Path -ne ""){
            $BookFiles = Get-Item -Path $Path
        } else {
            $BookFiles = $InputObject
        }

        # ブック毎に繰り返し
        foreach($BookFile in $BookFiles){
            # ブックを読み取り専用で開く
            $book = $excel.Workbooks.Open($BookFile, 0, $true)

            # 対象シートを取得
            $sheets = $excel.WorkSheets

            # シート毎に繰り返し
            foreach($sheet in $sheets){
                [PSCustomObject]@{ブック名 = $book.Name; シート名 = $sheet.Name} | Write-Output
            }
            
            # ブックを閉じる
            $book.Close(0)
        }

        # リソース解放
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }	
    end {
        # エクセルを終了
        $excel.Quit()

        # リソース解放
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
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