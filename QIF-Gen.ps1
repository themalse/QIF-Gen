$Host.UI.RawUI.BackgroundColor = 'black'
$revision = 'v0.62 Alpha'
$QIFFileName = 'test.QIF'
$StartDate = ''
$EndDate = ''
$TransactionsPerDay = ''
$TotalDays = ''

# VISUAL FORMATS #
$IsDefaultFG = (Get-Host).UI.RawUI.ForegroundColor
$IsDefaultBG = (Get-Host).UI.RawUI.BackgroundColor
$IsInvalidFG = 'red'
$IsInvalidBG = $IsDefaultBG
$lblInFocusFG = 'cyan'
$lblInFocusBG = $IsDefaultBG
#$InFocusFG = 'cyan'
#$InFocusBG = $IsDefaultBG
$lblIsSetFG = 'green'
$lblIsSetBG = $IsDefaultBG
$IsSetFG = 'green'
$IsSetBG = $IsDefaultBG


$lblStartDateFG = $IsDefaultFG
$lblStartDateBG = $IsDefaultFG
$StartDateFG = $IsDefaultFG
$StartDateBG = $IsDefaultBG

$lblEndDateFG = $IsDefaultFG
$lblEndDateBG = $IsDefaultFG
$EndDateFG = $IsDefaultFG
$EndDateBG = $IsDefaultBG

$lblTransactionsFG = $IsDefaultFG
$lblTransactionsBG = $IsDefaultFG
$TransactionsFG = $IsDefaultFG
$TransactionsBG = $IsDefaultBG

# VALIDITY CHECKS #
$StartDateValid = $false
$EndDateValid = $false
$TPerDayValid = $false
$IsConfirmed = $false
$AllValid = $false





function Test-Date($String) {
    try { [DateTime]::ParseExact($String,'dd/MM/yyyy',[System.Globalization.CultureInfo]'en-NZ')
        return $true
    }   catch 
        {
            return $false
        }
}

function Test-Int ([string]$x) {
    try {
        if ((0 + $x) -lt 1) {
            return $false
        }
        return $true
    }   catch {
            return $false
        }
}

function Update-UI {
    if ($AllValid -eq $true) {
        $lblStartDateFG = $lblIsSetFG
        $lblStartDateBG = $lblIsSetBG
        $StartDateFG = $IsSetFG
        $StartDateBG = $IsSetBG
        $lblEndDateFG = $lblIsSetFG
        $lblEndDateBG = $lblIsSetBG
        $EndDateFG = $IsSetFG
        $EndDateBG = $IsSetBG
        $lblTransactionsFG = $lblIsSetFG
        $lblTransactionsBG = $lblIsSetBG
        $TransactionsFG = $IsSetFG
        $TransactionsBG = $IsSetBG
    }   elseif ($StartDateValid -eq $false) {
            $lblStartDateFG = $lblInFocusFG
            $lblStartDateBG = $lblInFocusBG
            #$StartDateFG = $InFocusFG
            #$StartDateBG = $InFocusBG
            $lblEndDateFG = $IsDefaultFG
            $lblEndDateBG = $IsDefaultBG
            $EndDateFG = $IsDefaultFG
            $EndDateBG = $IsDefaultBG
            $lblTransactionsFG = $IsDefaultFG
            $lblTransactionsBG = $IsDefaultBG
            $TransactionsFG = $IsDefaultFG
            $TransactionsBG = $IsDefaultBG
        }   elseif ($EndDateValid -eq $false) {
                $lblStartDateFG = $lblIsSetFG
                $lblStartDateBG = $lblIsSetBG
                $StartDateFG = $IsSetFG
                $StartDateBG = $IsSetBG
                $lblEndDateFG = $lblInFocusFG
                $lblEndDateBG = $lblInFocusBG
                #$EndDateFG = $InFocusFG
                #$EndDateBG = $InFocusBG
                $lblTransactionsFG = $IsDefaultFG
                $lblTransactionsBG = $IsDefaultBG
                $TransactionsFG = $IsDefaultFG
                $TransactionsBG = $IsDefaultBG
            }   elseif ($TPerDayValid -eq $false) {
                    $lblStartDateFG = $lblIsSetFG
                    $lblStartDateBG = $lblIsSetBG
                    $StartDateFG = $IsSetFG
                    $StartDateBG = $IsSetBG
                    $lblEndDateFG = $lblIsSetFG
                    $lblEndDateBG = $lblIsSetBG
                    $EndDateFG = $IsSetFG
                    $EndDateBG = $IsSetBG
                    $lblTransactionsFG = $lblInFocusFG
                    $lblTransactionsBG = $lblInFocusBG
                    #$TransactionsFG = $InFocusFG
                    #$TransactionsBG = $InFocusBG
                }   elseif ($TPerDayValid -eq $true) {
                        $lblStartDateFG = $lblIsSetFG
                        $lblStartDateBG = $lblIsSetBG
                        $StartDateFG = $IsSetFG
                        $StartDateBG = $IsSetBG
                        $lblEndDateFG = $lblIsSetFG
                        $lblEndDateBG = $lblIsSetBG
                        $EndDateFG = $IsSetFG
                        $EndDateBG = $IsSetBG
                        $lblTransactionsFG = $IsSetFG
                        $lblTransactionsBG = $IsSetBG
                        $TransactionsFG = $IsSetFG
                        $TransactionsBG = $IsSetBG
                    }

    Write-Host "QIF Dummy Data Builder $revision"
    Write-Host
    Write-Host "Date Ranges to use for Dummy QIF file"
    Write-Host "Start Date: " -NoNewline -ForegroundColor $lblStartDateFG -BackgroundColor $lblStartDateBG
    try { $StartDate.ToString("dd/MM/yyyy") | Out-Null
        Write-Host $StartDate.ToString("dd/MM/yyyy") -ForegroundColor $StartDateFG -BackgroundColor $StartDateBG
    }
    catch {
        Write-Host $StartDate -ForegroundColor $StartDateFG -BackgroundColor $StartDateBG
    }
    Write-Host "End Date: " -NoNewline -ForegroundColor $lblEndDateFG -BackgroundColor $lblEndDateBG
    try { $EndDate.ToString("dd/MM/yyyy") | Out-Null
        Write-Host $EndDate.ToString("dd/MM/yyyy") -ForegroundColor $EndDateFG -BackgroundColor $EndDateBG
    }
    catch {
        Write-Host $EndDate -ForegroundColor $EndDateFG -BackgroundColor $EndDateBG
    }
    Write-Host
    Write-Host "Transactions/Day: " -NoNewline -ForegroundColor $lblTransactionsFG -BackgroundColor $lblTransactionsBG
    Write-Host $TransactionsPerDay -ForegroundColor $TransactionsFG -BackgroundColor $TransactionsBG
    Write-Host
    Write-Host "Total number of days to record: $TotalDays"
    try { Test-Int $TotalDays | Out-Null
        Write-Host 'Maxmimum possible number of transactions to generate:' ($TotalDays * $TransactionsPerDay) #ADDCOLOR
    }
    catch {
        Write-Host "Maxmimum possible number of transactions to generate: $TotalDays" #ADDCOLOR
    }
    Write-Host
}

Clear-Host
Write-Host

while ($AllValid -eq $false) {
    Update-UI
    if ($StartDateValid -eq $false) {
        $StartDate = Read-Host -Prompt 'Enter starting date (dd/MM/yyyy) for QIF dummy data' 
        Clear-Host
        if ($StartDate -eq '') {
            Write-Host 'Start Date' -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
            Write-Host " cannot be empty" -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
            $StartDateFG = $IsInvalidFG
            $StartDateBG = $IsInvalidBG
            }   else {
                    $StartDateValid = Test-Date $StartDate
                    if ($StartDateValid -eq $false) {
                        Write-Host $StartDate -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
                        Write-Host " is not a valid date format." -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
                        $StartDateFG = $IsInvalidFG
                        $StartDateBG = $IsInvalidBG
                    }   else {
                            $StartDate = $StartDateValid[0]
                            Write-Host
                        }
                }
    }

    elseif ($EndDateValid -eq $false) {
        $EndDate = Read-Host -Prompt 'Enter ending date (dd/MM/yyyy) for QIF dummy data'
        Clear-Host
        if ($EndDate -eq '') {
            Write-Host "End Date" -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
            Write-Host " cannot be empty" -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
            $EndDateFG = $IsInvalidFG
            $EndDateBG = $IsInvalidBG
        }   else {
                $EndDateValid = Test-Date $EndDate
                if ($EndDateValid -eq $false) {
                    Write-Host $EndDate -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
                    Write-Host " is not a valid date format." -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
                    $EndDateFG = $IsInvalidFG
                    $EndDateBG = $IsInvalidBG
                }   else {
                        $EndDate = $EndDateValid[0]
                        if ($StartDate -gt $EndDate) {
                            $EndDateValid = $false
                            Write-Host "End Date" -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
                            Write-Host " cannot be before the Start Date" -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
                            $EndDateFG = $IsInvalidFG
                            $EndDateBG = $IsInvalidBG
                        }   else {
                                $EndDate = $EndDateValid[0]
                                Clear-Host
                                Write-Host
                            }
                    }
            }
    }

    elseif ($TPerDayValid -eq $false) {
        $TransactionsPerDay = Read-Host "Enter the maximum possible number of transactions to generate per day"
        Clear-Host
        if ($TransactionsPerDay -eq '') {
            Write-Host "Transactions/Day" -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
            Write-Host " cannot be empty" -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
        }   else {
                $TPerDayValid = Test-Int $TransactionsPerDay
                if ($TPerDayValid -eq $false) {
                    $TransactionsFG = $IsInvalidFG
                    $TransactionsBG = $IsInvalidBG
                    Write-Host $TransactionsPerDay -NoNewline -ForegroundColor $IsInvalidFG -BackgroundColor $IsInvalidBG
                    Write-Host " must be an integer greater than or equal to 1" -ForegroundColor $IsDefaultFG -BackgroundColor $IsDefaultBG
                }   else {
                        $TransactionsFG = $IsSetFG
                        $TransactionsBG = $IsSetBG
                        $ts = New-TimeSpan -Start $StartDate -End $EndDate
                        $TotalDays = $ts.Days + 1
                        Clear-Host
                        Write-Host
                        Update-UI
                        $msg = 'Is the Transactions/Day value correct? [Y/N]' #ADDCOLOR
                        do {
                            $response = Read-Host -Prompt $msg
                            if ($response -eq 'n') {
                                $TPerDayValid = $false
                                $response = 'y'
                            }
                        } until ($response -eq 'y')
                        Clear-Host
                        Write-Host

                    }
            }
    }
    
    # TODO: Show actual number of transactions generated
    elseif ($IsConfirmed -eq $false) {
        $msg = 'Ready to create the QIF? [Y = Proceed | N = Quit]'
        do {
            $response = Read-Host -Prompt $msg
            if ($response -eq 'y') {
                $QIFContent = for ($i = 0; $i -lt $TotalDays; $i++) {
                    $jcount = Get-Random -Minimum 1 -Maximum $TransactionsPerDay
                    for ($j = 0; $j -lt $jcount; $j++) {
                        $tempnum = 0
                        $tempi = $StartDate.AddDays($i).ToString('dd/MM/yyyy')
                        "D$tempi"
                        while ($tempnum -eq 0) {
                            $tempnum = [math]::Round((Get-Random -Minimum -100000.00 -Maximum 100000.00),2)
                        }
                        "T$tempnum"
                        if ($tempnum -lt 0) {
                            "PWithdrawal"
                        }   else {
                                "PDeposit"
                        }
                        "M"
                        "^"
                    }
                }
                Out-File .\$QIFFileName -Encoding ascii
                Add-Content $QIFFileName '!Type:Bank'
                Add-Content $QIFFileName $QIFContent
            }
            Write-Host "The file has been saved to:"(Get-Location).Path"\"$QIFFileName
            $response = 'n'
        } until ($response -eq 'n')
        $AllValid = $true
    }
}
Read-Host -Prompt "Press Enter exit the tool"