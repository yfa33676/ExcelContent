# �o�̓I�u�W�F�N�g�̐ݒ�
$TypeData = @{TypeName = "ExcelContentValue"; DefaultDisplayPropertySet = "�u�b�N��", "�V�[�g��", "�Z���͈�", "�l"}
Update-TypeData @TypeData
$TypeData = @{TypeName = "ExcelSheetValue"; DefaultDisplayPropertySet = "�u�b�N��", "�V�[�g��"}
Update-TypeData @TypeData

function Get-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [System.IO.FileInfo]$FileObject, # ���̓I�u�W�F�N�g
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �擾�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$LiteralPath, # �t�@�C���p�X
        [string]$Value, # ����������
        [Alias("Name")] 
        [string]$SheetName, # �V�[�g��
        [string]$Range # �Z���͈�
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
        
        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
    }
    process {
        if($FileObject){
        } elseif($InputObject){
            $FileObject = $InputObject.FileObject
            $Value = $InputObject.�l
            $SheetName = $InputObject.�V�[�g��
            $Range = $InputObject.�Z���͈�
        } else{
            $FileObject = Get-Item -LiteralPath $LiteralPath
        }
        
        # �u�b�N��ǂݎ���p�ŊJ��
        $book = $excel.Workbooks.Open($FileObject, 0, $true)

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
                if($text -ne "" -and $text -match $Value){
                    [PSCustomObject]@{
                        PSTypeName = "ExcelContentValue"
                        FileObject = $FileObject
                        �u�b�N�� = $book.Name
                        �V�[�g�� = $sheet.Name
                        �Z���͈� = $cell.Address($false, $false)
                        �l = $text
                    } | Write-Output
                }
            }
        }
        
        # �t�@�C���I�u�W�F�N�g�̔j��
        $FileObject = $null
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(0)

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}

function Get-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [System.IO.FileInfo]$FileObject,
        [Parameter(Position = 0)]
        [string]$LiteralPath, # �t�@�C���p�X
        [Alias("Name")] 
        [string]$SheetName # �V�[�g��
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application
        
        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
        
    }
    process {
        if($FileObject){
            
        } else{
            $FileObject = Get-Item -LiteralPath $LiteralPath
        }
        
        # �u�b�N��ǂݎ���p�ŊJ��
        $book = $excel.Workbooks.Open($FileObject, 0, $true)

        # �ΏۃV�[�g���擾
        $sheets = $excel.WorkSheets | ? Name -match $SheetName

        # �V�[�g���ɌJ��Ԃ�
        foreach($sheet in $sheets){
            [PSCustomObject]@{
                PSTypeName = "ExcelSheetValue"
                FileObject = $FileObject
                �u�b�N�� = $book.Name
                �V�[�g�� = $sheet.Name
            } | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(0)

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}

function Set-ExcelContent{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [Alias("Name")] 
        [string]$SheetName, # �V�[�g��
        [string]$Range, # �Z���͈�
        [string]$Value # �ݒ�l
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application

        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
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

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}

function Add-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [Alias("Name")] 
        [string]$SheetName # �V�[�g��
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application

        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
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
            $sheet = $excel.WorkSheets.Add()
            try {
                $sheet.Name = [string]$InputObject.�V�[�g��
            } catch {
                $sheet.Delete()
                $_ | Write-Error
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

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}

function Delete-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [Alias("Name")] 
        [string]$SheetName # �V�[�g��
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application

        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
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

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}

function Set-ExcelSheet{
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$InputObject, # �ݒ�l�I�u�W�F�N�g
        [Parameter(Position = 0)]
        [string]$BookName, # �u�b�N��
        [Alias("Name")] 
        [string]$SheetName, # �V�[�g��
        [string]$Value # �ݒ�l
    )
    begin {
        # �G�N�Z�����N��
        $excel = New-Object -ComObject Excel.Application

        # �G�N�Z���ݒ�
        $excel.DisplayAlerts = $false
    }
    process {
        # �ݒ�l�I�u�W�F�N�g�𐶐�
        if($null -eq $InputObject){
            $InputObject = [PSCustomObject]@{�u�b�N�� = $BookName; �V�[�g�� = $SheetName; �l = $Value}
        } elseif($InputObject | Get-Member -Name �l){

        } else {
            $InputObject = [PSCustomObject]@{�u�b�N�� = $InputObject.�u�b�N��; �V�[�g�� = $InputObject.�V�[�g��; �l = $Value}
        }
        # �u�b�N�����擾
        $FullNames = Get-Item -Path $InputObject.�u�b�N��

        # �u�b�N�����ɌJ��Ԃ�
        foreach($FullName in $FullNames){
            # �u�b�N���J��
            $book = $excel.Workbooks.Open($FullName)
            
            # �ΏۃV�[�g��I��
            $sheet = $excel.WorkSheets.Item($InputObject.�V�[�g��)
            
            # �V�[�g����ύX
            $sheet.Name = [string]$InputObject.�l

            # �o�̓I�u�W�F�N�g�̃u�b�N����ҏW
            $InputObject.�u�b�N�� = $book.Name

            # �o��
            [PSCustomObject]$InputObject | Write-Output
        }
    }
    end {
        # �u�b�N�����
        $excel.WorkBooks | % Close(1)

        # �G�N�Z�����I��
        $excel.Quit()
        
        # NULL����
        $excel = $null
        
        # �K�x�[�W�R���N�V����
        [System.GC]::Collect()
    }
}
