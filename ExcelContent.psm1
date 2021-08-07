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
                        [PSCustomObject]@{�u�b�N�� = $book.Name; �V�[�g�� = $sheet.Name; �Z���Ԓn = $cell.Address($false, $false); �l = $text} | Write-Output
                    }
                }
            }
            
            # �u�b�N�����
            $book.Close(0)
        }

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }	
    end {
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
            
            # �u�b�N�����
            $book.Close(0)
        }

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }	
    end {
        # �G�N�Z�����I��
        $excel.Quit()

        # ���\�[�X���
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Set-ExcelContent{
  param(
      [Parameter(Position = 0, Mandatory)]
      [string]$ValuePath # �ݒ�lCSV�̃p�X
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
}