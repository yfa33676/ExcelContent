param(
    [Parameter(Position = 0)]
    [string[]]$Path = ".\*.xls*", # �t�@�C���p�X
    [string]$BookName, # �u�b�N��
    [string]$SheetName, # �V�[�g��
    [string]$Range, # �Z���͈�
    [string]$Pattern # ����������
)

if($Pattern -eq "" -and $BookName -eq "" -and $SheetName -eq "" -and $Range -eq ""){
  "������������͂��Ă�������"
  $BookName = Read-Host "�u�b�N��"
  $SheetName = Read-Host "�V�[�g��"
  $Range = Read-Host "�Z���͈�"
  $Pattern = Read-Host "����������"
}

# ����CSV
$result = New-Item ".\result.csv" -Type File -Force

# �Ώۃu�b�N���擾
$books = Get-Item -Path $Path
$books = $books | ? Name -match $BookName

# �w�b�_�[
$line = "�u�b�N��,�V�[�g��,�Z���Ԓn,�l"
$line | Write-Output
$line | Add-Content $result.FullName

# �G�N�Z�����N��
$excel = New-Object -ComObject Excel.Application
  

# �u�b�N���ɌJ��Ԃ�
for($i = 0; $i -lt $books.Count; $i++){
    # �u�b�N��ǂݎ���p�ŊJ��
    $book = $excel.Workbooks.Open($books[$i].FullName, 0, $true)

    # �ΏۃV�[�g���擾
    $sheets = $excel.WorkSheets | ? Name -match $SheetName

    # �V�[�g���ɌJ��Ԃ�
    foreach($sheet in $sheets){
        # �ΏۃZ���͈͂��擾
        if($Range -eq ""){
            $SelectedRange = $sheet.UsedRange
        } else {
            $SelectedRange = $sheet.Range("ZZ99,$Range")
        }

        # �ΏۃZ�����擾
        try{
            # �萔
            $constants = $SelectedRange.SpecialCells(
                [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeConstants, ## �萔
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlNumbers + ## ���l
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlTextValues ## �e�L�X�g
            )
        } catch{
            $constants = $null
        }
        try{
            # ����
            $formulas = $SelectedRange.SpecialCells(
                [Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeFormulas, ## ����
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlNumbers + ## ���l
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlTextValues + ## �e�L�X�g
                [Microsoft.Office.Interop.Excel.XlSpecialCellsValue]::xlLogical ## �_���l
            )
        } catch{
            $formulas = $null
        }
        # �ΏۃZ�� = �萔 + ����
        $cells = $constants + $formulas

        # �ΏۃZ�����ɌJ��Ԃ�
        foreach($cell in $cells){        
            # �Z���̃e�L�X�g���擾
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

# �G�N�Z�����I��
$excel.Quit()

# ���\�[�X���
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
