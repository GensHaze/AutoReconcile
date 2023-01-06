Param (
    [String] $NewCSV,
    [String] $ExistingCSV
)

Clear-Host
'a'
#region ----- Module loading
$ErrorActionPreference = 'Stop'
$scriptPath  = $PSScriptRoot
$modulePath  = "$scriptPath\Modules"
$archivePath = "$scriptPath\Archive"

Import-Module "$modulePath\Common.psm1" -Force -DisableNameChecking
#endregion


#region ----- Functions
Function Get-ConvertedDate {
    Param ([Parameter(ValueFromPipeline=$true)] $InputString)
    Switch ($InputString) {
        'ENE' { Return '01' }
        'FEB' { Return '02' }
        'MAR' { Return '03' }
        'ABR' { Return '04' }
        'MAY' { Return '05' }
        'JUN' { Return '06' }
        'JUL' { Return '07' }
        'AGO' { Return '08' }
        'SEP' { Return '09' }
        'OCT' { Return '10' }
        'NOV' { Return '11' }
        'DIC' { Return '12' } 
    }
}

Function Clear-SetupScreen {
    Param ([Int] $Count = 0, [Int] $Total = 0, [Switch] $SpecialColor)
    Clear-Host
    $colorSplat = @{}
    If ($SpecialColor -eq $True) {
        $colorSplat['ForegroundColor'] = 'Red'
        $colorSplat['BackgroundColor'] = 'DarkBlue'
    }
    If ($Count -ne 0 -or $Total -ne 0) {
        Write-Host "`r`nNew CSV: $referenceCSV"
        Write-Host "Old CSV: $differenceCSV`r`n"
        Write-Host "Transaction #$Count of $Total`r`n"
        Write-Host "Concept: $($record.Concepto)" @colorSplat
        Write-Host "Amount:  $($record.Amount)"   @colorSplat
        Write-Host "Type:    $recordClass`r`n"        @colorSplat
    }
}

Function Show-TransactionScreen {
    Param ([String] $ReplaceString)
    Clear-SetupScreen -Count $($a + 1) -Total $($addTable.Count)
    
    Invoke-Expression ($scriptAmountString -replace '<REPLACE>', $ReplaceString)
}

Function Show-ResultScreen {
    Param ([Switch] $Break, [String] $Table, [Switch] $CategoryPrompt)
    
    $specialColor = $False
    $colorSplat   = @{}
    If ($payee -eq 'Unknown' -or $category -eq 'Uncategorized' -or $subcategory -eq 'Uncategorized') {
        $specialColor = $True
        $colorSplat['SpecialColor'] = $True
    }
    Clear-SetupScreen -Count $($a + 1) -Total $($addTable.Count) -@colorSplat
    Write-Host "`r`nDetails:"
    $row = [PSCustomObject] @{
            Payee       = $payee
            Category    = $category
            Subcategory = $subcategory
            Beneficiary = $beneficiary
            Time        = $record.Time
            Account     = $account
            Amount      = $record.Amount
            Type        = $recordType
    }
    If   ($specialColor -eq $True) { $row | Out-String | Write-Host -ForegroundColor Red -BackgroundColor DarkBlue }
    Else { $row | Out-Host }

    If ($CategoryPrompt -eq $True) { $notes = Read-Host "Additional notes (enter '*' to change categories, '/' to start all over)" }
    Else { $notes = Read-Host "Additional notes (enter '*' to modify, '/' to start all over)" }
    Switch ($notes) {
        '*' {}
        '/' {
            $Global:a = -1
            Invoke-Expression ('[System.Collections.ArrayList] $<REPLACE> = @()' -replace '<REPLACE>', $Table)
            Return $True
        }
        Default {
            Add-Member -InputObject $row -MemberType NoteProperty -Name 'Notes' -Value $notes
            Invoke-Expression ('[Void] $<REPLACE>.Add($row)' -replace '<REPLACE>', $Table)
            Return $True
        }
    }
    If ($notes -ne '*') {
        
    }
}

$scriptAmountString = @'
    If ($record.Amount -lt 0) {
        $<REPLACE> = Get-ListItem -ItemType '<REPLACE>' -List $expense<REPLACE>List -Filename $expense<REPLACE>File 
        $expense<REPLACE>List = Get-Content $expense<REPLACE>File

        If ($<REPLACE> -ne '' -and $newPattern -ne '' -and $newPattern -notin $expense<REPLACE>Patterns.Pattern) {
            Try   { Add-Content -Path $expense<REPLACE>PatternsFile -Value "$newPattern,$<REPLACE>" }
            Catch { Write-Host "WARNING: Unable to write to file: $expense<REPLACE>PatternsFile" }
        }
    }
    Else {
        $<REPLACE> = Get-ListItem -ItemType '<REPLACE>' -List $income<REPLACE>List -Filename $income<REPLACE>File 
        $income<REPLACE>List = Get-Content $income<REPLACE>File

        If ($<REPLACE> -ne '' -and $newPattern -ne '' -and $newPattern -notin $income<REPLACE>Patterns.Pattern) {
            Try   { Add-Content -Path $income<REPLACE>PatternsFile -Value "$newPattern,$<REPLACE>" }
            Catch { Write-Host "WARNING: Unable to write to file: $income<REPLACE>PatternsFile" }
        }    
    }
    Return $<REPLACE>
'@

$scriptCompareString = @'
    Get-Variable -Include 'refCount*' | ForEach-Object { Set-Variable -Name $_.Name -Value 0 }
    ForEach ($referenceRecord in $<TABLE1>) {
        $referenceRecTime   = Get-Date $referenceRecord.Time -Format 'yyyyMMdd'
        $referenceRecAmount = "$([Math]::Abs($referenceRecord.Amount))" -replace '\.', '_'
        If ([Float] $referenceRecord.Amount -lt 0) { $refSign = 'm' } Else { $refSign = 'p' }
        $match = $False

        Invoke-Expression "`$refCount_${referenceRecTime}_${refSign}$referenceRecAmount++"
        $referenceCount = Invoke-Expression "`$refCount_${referenceRecTime}_${refSign}$referenceRecAmount"
        
        Get-Variable -Include 'difCount*' | ForEach-Object { Set-Variable -Name $_.Name -Value 0 }
        ForEach ($differenceRecord in $<TABLE2>) {
            $differenceRecTime   = Get-Date $differenceRecord.Time -Format 'yyyyMMdd'
            $differenceRecAmount = "$([Math]::Abs($differenceRecord.Amount))" -replace '\.', '_'
            If ([Float] $differenceRecord.Amount -lt 0) { $diffSign = 'm' } Else { $diffSign = 'p' }

            Invoke-Expression "`$difCount_${differenceRecTime}_${diffSign}$differenceRecAmount++"
            $differenceCount = Invoke-Expression "`$difCount_${differenceRecTime}_${diffSign}$differenceRecAmount"

            If ($referenceRecord.Amount -like $differenceRecord.Amount -and $referenceRecTime -like $differenceRecTime -and $referenceCount -le $differenceCount) { $match = $True ; Break }
        }
        If ($match -eq $False) { [Void] $<RESULTTABLE>.Add($referenceRecord) }
    }
'@
#endregion


#region ----- Variables
$PSDefaultParameterValues = @{
    'Get-Content:Encoding' = 'UTF8'
    'Import-CSV:Encoding'  = 'UTF8'
    'Export-CSV:Encoding'  = 'UTF8'
    'Add-Content:Encoding' = 'UTF8'
    'Out-File:Encoding'    = 'UTF8'
}

$accountListFile = "$modulePath\AccountList.txt"
$beneficiaryFile = "$modulePath\BeneficiaryList.txt"

$payeeListFile   = "$modulePath\PayeeList.csv"
$expenseCategoryFile = "$modulePath\ExpenseCategories.txt"
$incomeCategoryFile  = "$modulePath\IncomeCategories.txt"
$expenseSubcategoryFile = "$modulePath\ExpenseSubcategories.txt"
$incomeSubcategoryFile  = "$modulePath\IncomeSubcategories.txt"

$payeePatternsFile   = "$modulePath\PayeePatterns.csv"
$expenseCategoryPatternsFile = "$modulePath\ExpenseCategoriesPatterns.csv"
$incomeCategoryPatternsFile  = "$modulePath\IncomeCategoriesPatterns.csv"
$expenseSubcategoryPatternsFile = "$modulePath\ExpenseSubcategoriesPatterns.csv"
$incomeSubcategoryPatternsFile  = "$modulePath\IncomeSubcategoriesPatterns.csv"

$monthList = @(
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
)

$dateFormat    = "yyyy-MM-dd HH-mm-ss"
$paymentMatch  = 'SU PAGO|PAGO EN LINEA'
$interestMatch = 'INTERES|CREDISEGURO'

$recordInputMonth = 0
$year  = Get-Date -Format yyyy
$month = Get-Date -Format MM
$day   = Get-Date -Format dd
$today = Get-Date -Format 'dd/MM/yyyy'
$dataFilePath = "D:\Documents\Gens\Spreadsheets\Finances_$year\Data"
$dataFileTemplate = "Finances_${year}_${month}*.csv"
$financesFile = "D:\Documents\Gens\Spreadsheets\Finances_$year\Finances_$year.xlsx"

$optionsFile = [Ordered] @{
    '&Select a different file' = 'Open file browser'
    '&Continue without a file' = 'Do not reconcile with a file, process all transactions'
    '&Load this file' = "Load predefined file: $dataFileTemplate"
}
$optionsSkipFile = [Ordered] @{
    '&Select a file' = 'Open file browser'
    '&Continue without a file' = 'Do not reconcile with a file, process all transactions'
}
$optionsInput    = [Ordered] @{
    '&Load CSV'  = 'Load transaction data CSV'
    '&Input Manually' = 'Enter set number of transactions manually'
}

$compare    = $True
$updateList = $False
$requiredDifferenceColumns = @('Payee', 'Category', 'Subcategory', 'Time', 'Account', 'Amount', 'Type', 'Notes')

$columnsBanorte = @('Fecha', 'Concepto', 'Cargo')

[System.Collections.ArrayList] $addTable    = @()
[System.Collections.ArrayList] $removeTable = @()
[System.Collections.ArrayList] $resultAddTable    = @()
[System.Collections.ArrayList] $resultRemoveTable = @()
#endregion

$accountList = Get-Content $accountListFile
$account     = Get-ListItem -ItemType 'Account' -List $accountList -Filename $accountListFile
Write-Host "`r`nAccount set to: $account"

Switch ($account) {
    'Nomina CMS'     { $bank = 'Banorte - Nomina' }
    'Nomina Softtek' { $bank = 'Santander' }
    'Clasica'        { $bank = 'Banorte'   }
    'Los 40'         { $bank = 'Banorte'   }
    'Oro'            { $bank = 'Banorte'   }
    'Zero'           { $bank = 'HSBC'      }
    'Hey Banco'      { $bank = 'Hey'       }
    'RappiCard'      { $bank = 'Rappi'     }
    'RappiPay'       { $bank = 'Rappi'     }
    Default          { $bank = 'None'      }
}
Write-Host "Bank: $bank`r`n"

:Outer While ($True) {
    $inputConfirm = Read-Host '[1] [L]oad CSV transaction data, or [2] [I]nput manually?'

    Switch ($inputConfirm) {
        'L' { $inputProcess = 0 ; Break Outer }
        'I' { $inputProcess = 1 ; Break Outer }
        '1' { $inputProcess = 0 ; Break Outer }
        '2' { $inputProcess = 1 ; Break Outer }
        Default { Write-Host 'Invalid response, please try again' }
    }
}

If ($inputProcess -eq 1) {
    While ($True) {
        $transactionNumber = Read-Host 'Number of transactions to process'
        Try {
            $transactionNumber = [Convert]::ToInt32($transactionNumber)
            If ($transactionNumber -gt 200 -or $transactionNumber -le 0) { Throw }
            Break
        }
        Catch { Write-Host 'Invalid number, maximum is 200. Please try again.' }
    }
    $compare = $False
}
Else {
    $referenceCSV   = Open-FileBrowser -WindowTitle 'Select new transactions CSV' -InitialDirectory $scriptPath
    Switch ($bank) {
        'Banorte - Nomina' {
            $referenceTable = Import-CustomCSV -File $referenceCSV -WindowMessage "Select a different file" |
                Add-Member -MemberType ScriptProperty -Name 'Amount' -Value {[Double] $this.Abonos + [Double] $this.Cargos} -PassThru |
                Add-Member -MemberType ScriptProperty -Name 'Time' -Value {"$($this.Fecha.Substring(3,2))/$($this.Fecha.Substring(0,2))/$($this.Fecha.Substring(6))" | Get-Date} -PassThru #???? TODO: complicated to substring always, should be regex

        }
        'Banorte' {
            $referenceTable = Import-CustomCSV -File $referenceCSV -WindowMessage "Select a different file" -RequiredColumns $columnsBanorte |
                Add-Member -MemberType ScriptProperty -Name 'Amount' -Value {$this.Abono - $this.Cargo} -PassThru |
                Add-Member -MemberType ScriptProperty -Name 'Time' -Value {"$($this.Fecha.Substring(3,2))/$($this.Fecha.Substring(0,2))/$($this.Fecha.Substring(6))" | Get-Date} -PassThru #???? TODO: complicated to substring always, should be regex
        }
        'HSBC' {
            $referenceTable = Import-CustomCSV -File $referenceCSV -WindowMessage "Select a different file" -Headers @('Fecha', 'Concepto', 'Cargo') |
                Add-Member -MemberType ScriptProperty -Name 'Amount' -Value {-$this.Cargo} -PassThru |
                Add-Member -MemberType ScriptProperty -Name 'Time' -Value {"$($this.Fecha.Substring(3,3) | Get-ConvertedDate)/$($this.Fecha.Substring(0,2))/$year" | Get-Date} -PassThru
        }
        'Rappi' {
            $referenceTable = Import-CustomCSV -File $referenceCSV -WindowMessage "Select a different file" -Headers @('Fecha', 'Concepto', 'Moneda Extranjera', 'Cargo') |
                Add-Member -MemberType ScriptProperty -Name 'Amount' -Value {-$($this.Cargo -replace '\$', '')} -PassThru |
                Add-Member -MemberType ScriptProperty -Name 'Time' -Value {"$($this.Fecha.Substring(5,2))/$($this.Fecha.Substring(8,2))/$($this.Fecha.Substring(0,4))" | Get-Date} -PassThru
        }
        Default {
            $referenceTable = Import-CustomCSV -File $referenceCSV -WindowMessage "Select a different file" -RequiredColumns $columnsBanorte |
                Add-Member -MemberType ScriptProperty -Name 'Amount' -Value {$this.Abono - $this.Cargo} -PassThru |
                Add-Member -MemberType ScriptProperty -Name 'Time' -Value {"$($this.Fecha.Substring(3,2))/$($this.Fecha.Substring(0,2))/$($this.Fecha.Substring(6))" | Get-Date} -PassThru
        }
    }
    $referenceTable = $referenceTable | Sort-Object -Property Time, Concepto
    Write-Host "New transactions CSV set to: $referenceCSV`r`n"
}

#region ----- File select
Write-Host "`r`nWrite to file:"

$dataFiles = Get-ChildItem $dataFilePath -Filter "*.csv"
$dataInput = Get-ListItem -ItemType 'File' -List $dataFiles.Name -Default $month

If ($dataInput -eq $month) { $differenceCSV = "$dataFilePath\$($dataFiles.Name[$month - 1])" }
Else { $differenceCSV = "$dataFilePath\$dataInput" }
#endregion

#region ----- Compare
If ($compare -eq $True) {
    $differenceFullTable  = Import-CustomCSV -File $differenceCSV -WindowMessage "Select a different file" -RequiredColumns $requiredDifferenceColumns
    If   ($differenceFullTable.GetType().Name -like 'PSCustomObject') { [System.Collections.ArrayList] $differenceFullTable = @($differenceFullTable) }
    Else { [System.Collections.ArrayList] $differenceFullTable = $differenceFullTable }
    $differenceTable   = $differenceFullTable | Where-Object { $_.Account -like $account }
    
    Invoke-Expression ($scriptCompareString -replace '<TABLE1>', 'referenceTable'  -replace '<TABLE2>', 'differenceTable' -replace '<RESULTTABLE>', 'addTable')
    Invoke-Expression ($scriptCompareString -replace '<TABLE1>', 'differenceTable' -replace '<TABLE2>', 'referenceTable'  -replace '<RESULTTABLE>', 'removeTable')
}
ElseIf ($inputProcess -eq 1) {
    $emptyRow = [PSCustomObject] @{
        Payee       = ''
        Category    = ''
        Subcategory = ''
        Beneficiary = ''
        Time        = $today
        Account     = ''
        Amount      = ''
        Type        = ''
        Fecha       = $today
    }
    For ($i = 0 ; $i -lt $transactionNumber ; $i++) { $addTable.Add($emptyRow) }
}
Else { $addTable = $referenceTable }
#endregion

Clear-Host
Write-Host "Existing transactions CSV set to: $differenceCSV`r`n"
Write-Host "Number of new transactions to process: $($addTable.Count)"

If ($addTable.Count -eq $differenceTable.Count -and $inputProcess -ne 1) {
    Write-Host "WARNING: An equal amount of transactions to the reconcile file was detected, there might be an issue with matching, please investigate."
    Read-Host 'Press enter to continue'
}

#region ----- Process new records
$a = 0
While ($a -lt $addTable.Count) {
    $payeeList     = Import-CSV $payeeListFile
    $expenseCategoryList = Get-Content $expenseCategoryFile
    $incomeCategoryList  = Get-Content $incomeCategoryFile
    $expenseSubcategoryList = Get-Content $expenseSubcategoryFile
    $incomeSubcategoryList  = Get-Content $incomeSubcategoryFile

    $payeePatterns = Import-CSV $payeePatternsFile
    $expenseCategoryPatterns = Import-CSV $expenseCategoryPatternsFile
    $incomeCategoryPatterns  = Import-CSV $incomeCategoryPatternsFile
    $expenseSubCategoryPatterns = Import-CSV $expenseSubcategoryPatternsFile
    $incomeSubcategoryPatterns  = Import-CSV $incomeSubcategoryPatternsFile

    $beneficiaryList = Get-Content $beneficiaryFile

    $record = $addTable[$a]

    $payee    = 'Unknown'
    $category = 'Uncategorized'
    $subcategory = 'Uncategorized'
    $recordType  = 'Standard'

    If ($record.Concepto -cmatch $interestMatch) { $payee = $bank }
    Else {
        ForEach ($pattern in $payeePatterns) {
            If ($record.Concepto -cmatch $pattern.Pattern) { $payee = $pattern.Value ; Break }
        }
    }

    If ($record.Amount -lt 0) {
        ForEach ($pattern in $expenseCategoryPatterns)   {
            If ($record.Concepto -cmatch $pattern.Pattern) { $category = $pattern.Value ; Break }
        }
        ForEach ($pattern in $expenseSubcategoryPatterns) {
            If ($record.Concepto -cmatch $pattern.Pattern) { $subcategory = $pattern.Value ; Break }
        }
    }
    Else {
        ForEach ($pattern in $incomeCategoryPatterns)    {
            If ($record.Concepto -cmatch $pattern.Pattern) { $category = $pattern.Value ; Break }
        }
        ForEach ($pattern in $incomeSubcategoryPatterns) {
            If ($record.Concepto -cmatch $pattern.Pattern) { $subcategory = $pattern.Value ; Break }
        }
    }

    If ($record.Concepto -cmatch $paymentMatch)  { $recordType = 'Payment' }

    Clear-Host
    Write-Host "`r`nTransaction #$($a + 1) of $($addTable.Count)" #TODO: Add CSV filenames on this screen
    Write-Host "`r`nChoose beneficiary:"
    $beneficiary = Get-ListItem -ItemType 'Beneficiary' -List $beneficiaryList -Default 'All'
    

    If ($inputProcess -ne 1) {
        If ($record.Amount -lt 0) { $recordClass = 'Expense' } Else { $recordClass = 'Income' }
        $confirm = Show-ResultScreen -Table 'resultAddTable'

        If ($confirm -eq $True) { $a++ ; Continue }
        $newPattern = Read-Host 'Enter pattern to save (blank to skip)'
        $updateList = $True
    }
    Else {
        While ($True) {
            Clear-Host
            Write-Host "`r`nTransaction #$($a + 1) of $($addTable.Count)" #TODO: Add CSV filenames on this screen
            $record.Amount = Read-Host "`r`nEnter amount (negative numbers for expenses)"
            Try   {
                [Float] $record.Amount = [Math]::Round($record.Amount, 2)
                If ($record.Amount -lt 0) { $recordClass = 'Expense' } Else { $recordClass = 'Income' }
                Break
            }
            Catch { Write-Host 'Invalid input, please try again' }

        }
        While ($True) {
            Write-Host "`r`nDate input:"
            $recordInputMonth = Get-ListItem -ItemType 'Month' -List $monthList -ReturnIndex -Default $month
            $recordInputDay   = Read-Host "Enter day of the month (blank for default: $day)"
            If ($recordInputDay -eq '') { $recordInputDay = $day }
            Try {
                $record.Time = Get-Date -Year $year -Month $recordInputMonth -Day $recordInputDay -Hour 0 -Minute 0 -Second 0
                Break
            }
            Catch {
                Write-Host 'Invalid input, please try again'
                $recordInputMonth = 0
            }
        }
    }

    While ($True) {
        Clear-SetupScreen -Count $($a + 1) -Total $($addTable.Count)
        $payee = Get-ListItem -ItemType 'Payee' -List $payeeList.Name
        
        If ($payee -ne '' -and $newPattern -ne '' -and $newPattern -notin $payeePatterns.Pattern) {
            Try   { Add-Content -Path $payeePatternsFile -Value "$newPattern,$payee"  }
            Catch { Write-Host "WARNING: Unable to write to file: $payeePatternsFile" }
        }

        $payeeObject = $payeeList | Where-Object {$_.Name -eq $payee}
        $category    = $payeeObject.DefaultCategory
        $subcategory = $payeeObject.DefaultSubcategory

        $categoryConfirm = Show-ResultScreen -Table 'resultAddTable' -CategoryPrompt
        If ($categoryConfirm -eq $True) { $a++ ; Break }

        $category    = Show-TransactionScreen -ReplaceString 'Category'
        $subcategory = Show-TransactionScreen -ReplaceString 'Subcategory'

        While ($True) {
            If ($record.Amount -lt 0) { $classConfirm = Read-Host '[1] [S]tandard or [2] [D]eferred transaction?' }
            Else { $classConfirm = Read-Host '[1] [S]tandard or [2] [P]ayment transaction?' }
            If ($classConfirm.ToUpper() -notin @('S', 'D', 'P', '1', '2')) { Write-Host 'Invalid response, please try again.' }
            Else {
                Switch ($classConfirm.ToUpper()) {
                    'S' { $recordType = 'Standard' }
                    'D' { $recordType = 'Deferred' }
                    'P' { $recordType = 'Payment'  }
                    '1' { $recordType = 'Standard' }
                    '2' { If ($record.Amount -lt 0) { $recordType = 'Deferred' } Else { $recordType = 'Payment' } }
                }
                Break
            }
        }
        $confirm = Show-ResultScreen -Break -Table 'resultAddTable'
        
        If ($confirm -eq $True) {
            $a++
            If ($payee -notin $payeeList.Name) { Add-Content -Path $payeeListFile -Value "$payee,$category,$subcategory" }
            Break
        }
    }
}

Clear-Host
Write-Host "Number of old transactions to process: $($removeTable.Count)`r`n"

$a = 1
ForEach ($record in $removeTable) {
    Clear-SetupScreen -Count $a -Total $($removeTable.Count)
    $record | Out-Host

    $discardConfirm = Read-Host "Enter 'K' or 'k' to keep, anything else to discard"
    If ($discardConfirm.ToUpper() -ne 'K') {
        [Void] $differenceFullTable.Remove($record)
        [Void] $resultRemoveTable.Add($record)
    }
    $a++
}

Clear-SetupScreen

If ($addTable.Count -eq 0 -and $removeTable.Count -eq 0) {
    Write-Host 'No records to process.'
    Read-Host 'Press enter to exit'
    Exit 0
}
Else {
    $payeeList              | Sort-Object -Property Name | Export-CSV $payeeListFile -NoTypeInformation
    $expenseCategoryList    | Sort-Object | Out-File $expenseCategoryFile
    $incomeCategoryList     | Sort-Object | Out-File $incomeCategoryFile
    $expenseSubcategoryList | Sort-Object | Out-File $expenseSubcategoryFile
    $incomeSubcategoryList  | Sort-Object | Out-File $incomeSubcategoryFile
}

$timestamp = Get-Date -Format $dateFormat

Write-Host "`r`nAll transactions processed."
Write-Host "Exporting results to CSV: $differenceCSV"

While ($True) {
    Try {
        $differenceCSVArchiveName = "$($differenceCSV.Name)_$timestamp.csv"
        Copy-Item $differenceCSV -Destination "$archivePath\$differenceCSVArchiveName"
        If ($compare -eq $True) { $differenceFullTable | Sort-Object -Property Time, Payee | Export-CSV $differenceCSV -NoTypeInformation -Force }
        $resultAddTable      | Sort-Object -Property Time, Payee | Export-CSV $differenceCSV -NoTypeInformation -Append

        Break
    }
    Catch { Read-Host "$_`r`nPress enter to try again" }
}
$openConfirm = Read-Host "`r`nAll done, enter 'O' to open main file, anything else to exit"
If ($openConfirm.ToUpper() -eq 'O') { Invoke-Item $financesFile }

Trap {
    Write-Host "Terminating expression found - unable to continue, attempting export..."
    $_
    If ($resultAddTable.Count -gt 0) {
        Try {
            $timestamp  = Get-Date -Format $dateFormat
            $resultFile = "$scriptPath\result_$timestamp.csv"
            $resultAddTable | Sort-Object -Property Time | Export-CSV $resultFile -NoTypeInformation
        }
        Catch { Write-Host 'Unable to export data in exportTable, please retry.' }
    }
    If ($removeTable.Count -gt 0) {
        Try {
            $timestamp  = Get-Date -Format $dateFormat
            $reviewFile = "$scriptPath\review_$timestamp.csv"
            $removeTable | Sort-Object -Property Time | Export-CSV $reviewFile -NoTypeInformation
        }
        Catch { Write-Host 'Unable to export data in reviewTable, please retry.' }
    }
    Read-Host 'Enter to exit'
}