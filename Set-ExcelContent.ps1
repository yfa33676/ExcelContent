param(
    [Parameter(Position = 0)]
    [string]$ValuePath = ".\result.csv"
)
# �Ώۃu�b�N���擾
$books = Get-Content -Path $ValuePath | ConvertFrom-Csv | Select-Object �u�b�N�� -Unique | % �u�b�N�� | Get-Item

# �G�N�Z�����N��
$excel = New-Object -ComObject Excel.Application

# �u�b�N���ɌJ��Ԃ�
for($i = 0; $i -lt $books.Count; $i++){
    # �u�b�N���J��
    $book = $excel.Workbooks.Open($books[$i].FullName, 0, $false)
    
    # �ΏۃZ�����擾
    $cells = Get-Content -Path $ValuePath | ConvertFrom-Csv | Where-Object �u�b�N�� -eq $book.Name

    # �Z�����ɌJ��Ԃ�
    foreach($cell in $cells){
        $cell.�u�b�N�� + "," + $cell.�V�[�g�� + "," + $cell.�Z���Ԓn + "," + $cell.�l
        # �ΏۃV�[�g��I��
        $sheet = $excel.WorkSheets.Item($cell.�V�[�g��)
        # �ΏۃZ���ɒl��ݒ�
        $sheet.Cells.Range($cell.�Z���Ԓn).Value2 = $cell.�l -Replace "``n","`n"
    }

    # �u�b�N��ۑ����ĕ���
    $book.Close(1)
}
# �G�N�Z�����I��
$excel.Quit()

# ���\�[�X���
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null