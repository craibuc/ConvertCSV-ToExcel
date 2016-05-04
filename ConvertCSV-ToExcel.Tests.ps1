$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "ConvertCSV-ToExcel" {

    #
    # arrange
    #

    # dummy data
    $nl = [Environment]::NewLine
    $content = '"DATE_COLUMN","DATETIME_COLUMN","TEXT_COLUMN"' + $nl + '2015-05-01,2015-05-01 23:00:00.000,LOREM IPSUM' + $nl + 'NULL,null,nuLL'

    # create first CSV file and popoulate it with data
    $CSV_0 = New-item "TestDrive:\CSV_0.csv" -Type File
    Set-Content $CSV_0 -Value $content

    # create second CSV file and popoulate it with data
    $CSV_1 = New-item "TestDrive:\CSV_1.csv" -Type File
    Set-Content $CSV_1 -Value $content

    Context "Supplying the manditory parameters; -inputfile via the pipeline" {

        # act
        Get-ChildItem 'TestDrive:\*.csv' | ConvertCSV-ToExcel -Output 'Output.xlsx'

        It "Coonverts more than one CSV-formatted file to a single XLSX-formatted file in the current directory" {
            Get-ChildItem 'Output.xlsx' | Should Exist
        }

    }

}
