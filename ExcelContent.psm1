function Get-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [psobject]$InputObject, # ���̓I�u�W�F�N�g
        [Parameter(ValueFromPipeline, Position = 0)]
        [string]$Path, # �t�@�C���p�X
        [string]$Pattern, # ����������
        [string]$BookName, # �u�b�N��
        [string]$SheetName, # �V�[�g��
        [string]$Range # �Z���͈�
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �Ώۃu�b�N�̎擾
        if($Path -ne ""){
            $BookFiles = Get-Item -Path $Path | ? Name -match $BookName
        } else {
            $BookFiles = $InputObject | ? Name -match $BookName
        }

        # �u�b�N���ɌJ��Ԃ�
        foreach($BookFile in $BookFiles){
            # �u�b�N��ǂݎ���p�ŊJ��
            $book = $excel.Workbooks.Open($BookFile, 0, $true)

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
                        [PSCustomObject]@{�u�b�N�� = $book.Name; �V�[�g�� = $sheet.Name; �Z���͈� = $cell.Address($false, $false); �l = $text} | Write-Output
                    }
                }
            }
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(0)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
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
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �Ώۃu�b�N�̎擾
        if($Path -ne ""){
            $BookFiles = Get-Item -Path $Path
        } else {
            $BookFiles = $InputObject
        }

        # �u�b�N���ɌJ��Ԃ�
        foreach($BookFile in $BookFiles){
            # �u�b�N��ǂݎ���p�ŊJ��
            $book = $excel.Workbooks.Open($BookFile, 0, $true)

            # �ΏۃV�[�g���擾
            $sheets = $excel.WorkSheets

            # �V�[�g���ɌJ��Ԃ�
            foreach($sheet in $sheets){
                [PSCustomObject]@{�u�b�N�� = $book.Name; �V�[�g�� = $sheet.Name} | Write-Output
            }
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(0)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Set-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [string]$SheetName, # �V�[�g��
        [string]$Range, # �Z���͈�
        [string]$Value # �ݒ�l
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �ݒ�l�I�u�W�F�N�g�𐶐�
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{�u�b�N�� = $BookName; �V�[�g�� = $SheetName; �Z���͈� = $Range; �l = $Value}
        }
        # �u�b�N�����擾
        $FullNames = Get-Item -Path $InputObject.�u�b�N�� | % FullName

        # �u�b�N�����ɌJ��Ԃ�
        foreach($FullName in $FullNames){
            # �u�b�N���J��
            $book = $excel.Workbooks.Open($FullName)
            
            # �ΏۃV�[�g��I��
            $sheet = $excel.WorkSheets.Item($InputObject.�V�[�g��)
            
            # �ΏۃZ���ɒl��ݒ�
            $sheet.Cells.Range($InputObject.�Z���͈�).Value2 = $InputObject.�l -Replace "``n","`n"

            # �o�̓I�u�W�F�N�g�̃u�b�N����ҏW
            $InputObject.�u�b�N�� = $book.Name

            # �o��
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(1)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Add-ExcelSheets{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [string]$SheetName # �V�[�g��
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �ݒ�l�I�u�W�F�N�g�𐶐�
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{�u�b�N�� = $BookName; �V�[�g�� = $SheetName}
        }
        # �u�b�N�����擾
        $FullNames = Get-Item -Path $InputObject.�u�b�N�� | % FullName

        # �u�b�N�����ɌJ��Ԃ�
        foreach($FullName in $FullNames){
            # �u�b�N���J��
            $book = $excel.Workbooks.Open($FullName)
            
            # �ΏۃV�[�g��ǉ�
            $sheet = $book.WorkSheets.Add()
            try {
                $sheet.Name = $InputObject.�V�[�g��
            } catch {
                $_ | Write-Error
                $sheet.Delete()
            }

            # �o�̓I�u�W�F�N�g�̃u�b�N����ҏW
            $InputObject.�u�b�N�� = $book.Name

            # �o��
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(1)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Delete-ExcelSheets{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [string]$SheetName # �V�[�g��
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �ݒ�l�I�u�W�F�N�g�𐶐�
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{�u�b�N�� = $BookName; �V�[�g�� = $SheetName}
        }
        # �u�b�N�����擾
        $FullNames = Get-Item -Path $InputObject.�u�b�N�� | % FullName

        # �u�b�N�����ɌJ��Ԃ�
        foreach($FullName in $FullNames){
            # �u�b�N���J��
            $book = $excel.Workbooks.Open($FullName)
            
            # �ΏۃV�[�g��I��
            $sheet = $excel.WorkSheets.Item($InputObject.�V�[�g��)
            
            # �V�[�g���폜
            $sheet.Delete()

            # �o�̓I�u�W�F�N�g�̃u�b�N����ҏW
            $InputObject.�u�b�N�� = $book.Name

            # �o��
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(1)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Set-ExcelSheets{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [string]$SheetName, # �V�[�g��
        [string]$Value # �ݒ�l
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
    }
    process {
        # �ݒ�l�I�u�W�F�N�g�𐶐�
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{�u�b�N�� = $BookName; �V�[�g�� = $SheetName; �l = $Value}
        }
        # �u�b�N�����擾
        $FullNames = Get-Item -Path $InputObject.�u�b�N�� | % FullName

        # �u�b�N�����ɌJ��Ԃ�
        foreach($FullName in $FullNames){
            # �u�b�N���J��
            $book = $excel.Workbooks.Open($FullName)
            
            # �ΏۃV�[�g��I��
            $sheet = $excel.WorkSheets.Item($InputObject.�V�[�g��)
            
            # �V�[�g����ύX
            $sheet.Name = $InputObject.�l

            # �o�̓I�u�W�F�N�g�̃u�b�N����ҏW
            $InputObject.�u�b�N�� = $book.Name

            # �o��
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(1)

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null

        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
