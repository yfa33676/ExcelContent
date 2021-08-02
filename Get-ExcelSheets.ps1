param(
    [Parameter(Position = 0)]
    [string[]]$Path = ".\*.xls*", # �t�@�C���p�X
    [string]$BookName = "", # �u�b�N��
    [string]$SheetName = "" # �V�[�g��
)

if($BookName -eq "" -and $SheetName -eq ""){
  "������������͂��Ă�������"
  $BookName = Read-Host "�u�b�N��"
  $SheetName = Read-Host "�V�[�g��"
}

# ����CSV
$result = New-Item ".\result.csv" -Type File -Force

# �Ώۃu�b�N���擾
$books = Get-Item -Path $Path
$books = $books | ? Name -match $BookName

# �w�b�_�[
$line = "�u�b�N��,�V�[�g��"
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
        $line = $book.Name + "," + $sheet.Name
        $line | Write-Output
        $line | Add-Content $result.FullName
    }
    $book.Close(0)
}

# �G�N�Z�����I��
$excel.Quit()

# ���\�[�X���
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
