# 出力オブジェクトの設定
$TypeData = @{TypeName = "ExcelContentValue"; DefaultDisplayPropertySet = "ブック名", "シート名", "セル範囲", "値"}
Update-TypeData @TypeData
$TypeData = @{TypeName = "ExcelSheetValue"; DefaultDisplayPropertySet = "ブック名", "シート名"}
Update-TypeData @TypeData

function Get-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [System.IO.FileInfo]$FileObject, # 入力オブジェクト
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # 取得値オブジェクト
        [Parameter(Position = 0)]
        [string]$LiteralPath, # ファイルパス
        [string]$Value, # 検索文字列
        [Alias("Name")] 
        [string]$SheetName, # シート名
        [string]$Range # セル範囲
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application
        
        # エクセル設定
        $excel.DisplayAlerts = $false
    }
    process {
        if($FileObject){
        } elseif($InputObject){
            $FileObject = $InputObject.FileObject
            $Value = $InputObject.値
            $SheetName = $InputObject.シート名
            $Range = $InputObject.セル範囲
        } else{
            $FileObject = Get-Item -LiteralPath $LiteralPath
        }
        
        # ブックを読み取り専用で開く
        $book = $excel.Workbooks.Open($FileObject, 0, $true)

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
                if($text -ne "" -and $text -match $Value){
                    [PSCustomObject]@{
                        PSTypeName = "ExcelContentValue"
                        FileObject = $FileObject
                        ブック名 = $book.Name
                        シート名 = $sheet.Name
                        セル範囲 = $cell.Address($false, $false)
                        値 = $text
                    } | Write-Output
                }
            }
        }
        
        # ファイルオブジェクトの破棄
        $FileObject = $null
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(0)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}

function Get-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [System.IO.FileInfo]$FileObject,
        [Parameter(Position = 0)]
        [string]$LiteralPath, # ファイルパス
        [Alias("Name")] 
        [string]$SheetName # シート名
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application
        
        # エクセル設定
        $excel.DisplayAlerts = $false
        
    }
    process {
        if($FileObject){
            
        } else{
            $FileObject = Get-Item -LiteralPath $LiteralPath
        }
        
        # ブックを読み取り専用で開く
        $book = $excel.Workbooks.Open($FileObject, 0, $true)

        # 対象シートを取得
        $sheets = $excel.WorkSheets | ? Name -match $SheetName

        # シート毎に繰り返し
        foreach($sheet in $sheets){
            [PSCustomObject]@{
                PSTypeName = "ExcelSheetValue"
                FileObject = $FileObject
                ブック名 = $book.Name
                シート名 = $sheet.Name
            } | Write-Output
        }
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(0)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}

function Set-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # 設定値オブジェクト
        [Parameter(Position = 0)]
        [string]$BookName, # ブック名
        [Alias("Name")] 
        [string]$SheetName, # シート名
        [string]$Range, # セル範囲
        [string]$Value # 設定値
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application

        # エクセル設定
        $excel.DisplayAlerts = $false
    }
    process {
        # 設定値オブジェクトを生成
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{ブック名 = $BookName; シート名 = $SheetName; セル範囲 = $Range; 値 = $Value}
        }
        # ブック名を取得
        $FullNames = Get-Item -Path $InputObject.ブック名 | % FullName

        # ブック名毎に繰り返し
        foreach($FullName in $FullNames){
            # ブックを開く
            $book = $excel.Workbooks.Open($FullName)
            
            # 対象シートを選択
            $sheet = $excel.WorkSheets.Item($InputObject.シート名)
            
            # 対象セルに値を設定
            $sheet.Cells.Range($InputObject.セル範囲).Value2 = $InputObject.値 -Replace "``n","`n"

            # 出力オブジェクトのブック名を編集
            $InputObject.ブック名 = $book.Name

            # 出力
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(1)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}

function Add-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # 設定値オブジェクト
        [Parameter(Position = 0)]
        [string]$BookName, # ブック名
        [Alias("Name")] 
        [string]$SheetName # シート名
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application

        # エクセル設定
        $excel.DisplayAlerts = $false
    }
    process {
        # 設定値オブジェクトを生成
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{ブック名 = $BookName; シート名 = $SheetName}
        }
        # ブック名を取得
        $FullNames = Get-Item -Path $InputObject.ブック名 | % FullName

        # ブック名毎に繰り返し
        foreach($FullName in $FullNames){
            # ブックを開く
            $book = $excel.Workbooks.Open($FullName)
            
            # 対象シートを追加
            $sheet = $excel.WorkSheets.Add()
            try {
                $sheet.Name = [string]$InputObject.シート名
            } catch {
                $sheet.Delete()
                $_ | Write-Error
            }

            # 出力オブジェクトのブック名を編集
            $InputObject.ブック名 = $book.Name

            # 出力
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(1)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}

function Delete-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # 設定値オブジェクト
        [Parameter(Position = 0)]
        [string]$BookName, # ブック名
        [Alias("Name")] 
        [string]$SheetName # シート名
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application

        # エクセル設定
        $excel.DisplayAlerts = $false
    }
    process {
        # 設定値オブジェクトを生成
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{ブック名 = $BookName; シート名 = $SheetName}
        }
        # ブック名を取得
        $FullNames = Get-Item -Path $InputObject.ブック名 | % FullName

        # ブック名毎に繰り返し
        foreach($FullName in $FullNames){
            # ブックを開く
            $book = $excel.Workbooks.Open($FullName)
            
            # 対象シートを選択
            $sheet = $excel.WorkSheets.Item($InputObject.シート名)
            
            # シートを削除
            $sheet.Delete()

            # 出力オブジェクトのブック名を編集
            $InputObject.ブック名 = $book.Name

            # 出力
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(1)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}

function Set-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # 設定値オブジェクト
        [Parameter(Position = 0)]
        [string]$BookName, # ブック名
        [Alias("Name")] 
        [string]$SheetName, # シート名
        [string]$Value # 設定値
    )
    begin {
        # エクセルを起動
        $excel = New-Object -ComObject Excel.Application

        # エクセル設定
        $excel.DisplayAlerts = $false
    }
    process {
        # 設定値オブジェクトを生成
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{ブック名 = $BookName; シート名 = $SheetName; 値 = $Value}
        } elseif($InputObject | Get-Member -Name 値){

        } else {
            $InputObject = [PSCustomObject]@{ブック名 = $InputObject.ブック名; シート名 = $InputObject.シート名; 値 = $Value}
        }
        # ブック名を取得
        $FullNames = Get-Item -Path $InputObject.ブック名

        # ブック名毎に繰り返し
        foreach($FullName in $FullNames){
            # ブックを開く
            $book = $excel.Workbooks.Open($FullName)
            
            # 対象シートを選択
            $sheet = $excel.WorkSheets.Item($InputObject.シート名)
            
            # シート名を変更
            $sheet.Name = [string]$InputObject.値

            # 出力オブジェクトのブック名を編集
            $InputObject.ブック名 = $book.Name

            # 出力
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # ブックを閉じる
        $excel.WorkBooks | % Close(1)

        # エクセルを終了
        $excel.Quit()
        
        # NULLを代入
        $excel = $null
        
        # ガベージコレクション
        [System.GC]::Collect()
    }
}
