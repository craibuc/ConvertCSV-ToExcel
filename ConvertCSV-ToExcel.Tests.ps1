$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "ConvertCSV-ToExcel" {

    #
    # arrange
    #

    $nl = [Environment]::NewLine
    $xlsx` = 'Output.xlsx'

    # dummy data (HERE string)
$Content = @'
"DATE_COLUMN","DATETIME_COLUMN","TEXT_COLUMN"
"2015-05-01","2015-05-01 23:00:00.000","LOREM IPSUM"
"2015-05-01","2015-05-01 23:00:00.000","LOREM IPSUM"
'@

    # create first CSV file and popoulate it with data
    $CSV_0 = New-item "TestDrive:\CSV_0.csv" -Type File
    Set-Content $CSV_0 -Value $content

    # create second CSV file and popoulate it with data
    $CSV_1 = New-item "TestDrive:\CSV_1.csv" -Type File
    Set-Content $CSV_1 -Value $content

    Context "Supplying the manditory parameters; -InputFile via the pipeline" {

        It -"Coonverts more than one CSV-formatted file to a single XLSX-formatted file in the current directory" {

            # make TestDrive:\ the current directory
            pushd 'TestDrive:'

            # act
            Get-ChildItem 'TestDrive:\*.csv' | ConvertCSV-ToExcel -Output $xlsx -Verbose

            # assert
            Get-ChildItem ".\$xlsx" | Should Exist

            # restore current directory
            popd

        }

    }

    Context "Supplying the -Path parameter" {

        It -skip "Converts more than one CSV-formatted file to a single XLSX-formatted file in the specified directory" {
            # act
            Get-ChildItem 'TestDrive:\*.csv' | ConvertCSV-ToExcel -Output $xlsx -Path 'TestDrive:' -Verbose

            # assert
            Get-ChildItem "TestDrive:\$xlsx" | Should Exist
        }

    }

    Context "Supplying the -EnableAutoFilter parameter" {
        It "Enables the auto-filter for each sheet in the workbook" {}
    }

    Context "Supplying the -FreezePanes parameter" {
        It "Freezes the top row of each sheet in the workbook" {}
    }

    Context "Supplying the -BoldHeader parameter" {
        It "Bolds the top row of each sheet in the workbook" {}
    }

}
