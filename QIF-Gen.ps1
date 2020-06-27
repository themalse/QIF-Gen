$revision = 'v0.7'

# Clear variables if used Globally
$StartDate = ''
$EndDate = ''
$TransactionsPerDay = ''
$TotalDays = ''
$QIFFileName = 'test.QIF' #File name to use as default

# UI colour formatting | [Forground Label,Background Label,Foreground Val,Background Val]
$DefaultCol = 'gray','black','gray','black'
$ValidCol = 'green',$DefaultCol[1],'green',$DefaultCol[1]
$InvalidCol = 'red',$DefaultCol[1],'red',$DefaultCol[1]
$InFocusCol = 'cyan',$DefaultCol[1],'cyan',$DefaultCol[1]

# Set default UI colours
$Host.UI.RawUI.ForegroundColor = $DefaultCol[0]
$Host.UI.RawUI.BackgroundColor = $DefaultCol[1]
$StartDateCol = $DefaultCol[0],$DefaultCol[1],$DefaultCol[0],$DefaultCol[1]
$EndDateCol = $DefaultCol[0],$DefaultCol[1],$DefaultCol[0],$DefaultCol[1]
$TransactionsCol = $DefaultCol[0],$DefaultCol[1],$DefaultCol[0],$DefaultCol[1]

# Validity Checks
$StartDateValid = $null,$false
$EndDateValid = $null,$false
$TransactionsValid = $false
$AllValid = $false

function Test-Date($String) {
    try { [DateTime]::ParseExact($String,'dd/MM/yyyy',[System.Globalization.CultureInfo]'en-NZ')
        return $true
    }   catch 
        {
            return $null,$false
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
        $StartDateCol = $ValidCol
        $EndDateCol = $ValidCol
        $TransactionsCol = $ValidCol
    }   elseif ($StartDateValid[1] -eq $false) {
            $StartDateCol = $InFocusCol
            $EndDateCol = $DefaultCol
            $TransactionsCol = $DefaultCol
        }   elseif ($EndDateValid[1] -eq $false) {
                $StartDateCol = $ValidCol
                $EndDateCol = $InFocusCol
                $TransactionsCol = $DefaultCol
            }   elseif ($TransactionsValid -eq $false) {
                    $StartDateCol = $ValidCol
                    $EndDateCol = $ValidCol
                    $TransactionsCol = $InFocusCol
                }   elseif ($TransactionsValid -eq $true) {
                    $StartDateCol = $ValidCol
                    $EndDateCol = $ValidCol
                    $TransactionsCol = $ValidCol
                    }
    Clear-Host
    Write-Host 'QIF-Gen'$revision': Builds a QIF file with dummy data'
    Write-Host
    Write-Host "Date Ranges to use for Dummy QIF file"
    Write-Host "Start Date: " -NoNewline -ForegroundColor $StartDateCol[0] -BackgroundColor $StartDateCol[1]
    if ($StartDateValid[1] -eq $false) {
        Write-Host $StartDate -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
    }   else {
        Write-Host $StartDate.ToString("dd/MM/yyyy") -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
        }
    Write-Host "End Date: " -NoNewline -ForegroundColor $EndDateCol[0] -BackgroundColor $EndDateCol[1]
    if ($EndDateValid[1] -eq $false) {
        Write-Host $EndDate -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
    }   else {
            Write-Host $EndDate.ToString("dd/MM/yyyy") -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
        }
    Write-Host
    Write-Host "Transactions/Day: " -NoNewline -ForegroundColor $TransactionsCol[0] -BackgroundColor $TransactionsCol[1]
    Write-Host $TransactionsPerDay -ForegroundColor $TransactionsCol[2] -BackgroundColor $TransactionsCol[3]
    Write-Host
    Write-Host "Total number of days: $TotalDays"
    try { Test-Int $TotalDays | Out-Null
        Write-Host 'Maxmimum possible number of transactions to generate:' ($TotalDays * $TransactionsPerDay) #ADDCOLOR
    }
    catch {
        Write-Host "Maxmimum possible number of transactions to generate: $TotalDays" #ADDCOLOR
    }
    Write-Host
}

Clear-Host #Clears terminal
Update-UI

while ($AllValid -eq $false) {
    if ($StartDateValid[1] -eq $false) {
        $StartDate = Read-Host -Prompt 'Enter starting date (dd/MM/yyyy)' 
        if ($StartDate -eq '') {
            Update-UI
            Write-Host 'Start Date' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
            Write-Host " cannot be empty" -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
        }   else {
                $StartDateValid = Test-Date $StartDate
                if ($StartDateValid[1] -eq $false) {
                    Update-UI
                    Write-Host $StartDate -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                    Write-Host " is not a valid date format." -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
                }   else {
                        $StartDate = $StartDateValid[0]
                        Update-UI
                    }
            }
    }

    elseif ($EndDateValid[1] -eq $false) {
        $EndDate = Read-Host -Prompt 'Enter ending date (dd/MM/yyyy)'
        if ($EndDate -eq '') {
            Update-UI
            Write-Host 'End Date' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
            Write-Host ' cannot be empty' -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
        }   else {
                $EndDateValid = Test-Date $EndDate
                if ($EndDateValid[1] -eq $false) {
                    Update-UI
                    Write-Host $EndDate -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                    Write-Host ' is not a valid date format.' -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
                }   else {
                        $EndDate = $EndDateValid[0]
                        if ($StartDate -gt $EndDate) {
                            $EndDateValid[1] = $false
                            Write-Host 'End Date' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                            Write-Host ' cannot be before the Start Date' -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
                        }   else {
                                $EndDate = $EndDateValid[0]
                                Update-UI
                            }
                    }
            }
    }

    elseif ($TransactionsValid -eq $false) {
        $TransactionsPerDay = Read-Host 'Enter maximum random transactions per day'
        if ($TransactionsPerDay -eq '') {
            Update-UI
            Write-Host "Transactions/Day" -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
            Write-Host " cannot be empty" -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
        }   else {
                $TransactionsValid = Test-Int $TransactionsPerDay
                if ($TransactionsValid -eq $false) {
                    Update-UI
                    Write-Host $TransactionsPerDay -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                    Write-Host " must be an integer greater than or equal to 1" -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
                }   else {
                        $TotalDays = (New-TimeSpan -Start $StartDate -End $EndDate).Days + 1
                        Update-UI
                        $msg = 'Are Transactions/Day value correct? [Y|N]' #ADDCOLOR
                        do {
                            $response = Read-Host -Prompt $msg
                            if ($response -eq 'n') {
                                $TransactionsValid = $false
                                $response = 'y'
                            }
                        } until ($response -eq 'y')
                        Update-UI
                    }
            }
    }

    elseif (($StartDateValid -eq $true) -and ($EndDateValid -eq $true) -and ($TransactionsValid -eq $true)) {
        do {
            $response = Read-Host -Prompt 'Create QIF? [Y|N]'
            if ($response -eq 'y') {
                $QIFContent =   for ($i = 0; $i -lt $TotalDays; $i++) {
                                    $jcount = Get-Random -Minimum 1 -Maximum $TransactionsPerDay
                                    for ($j = 0; $j -lt $jcount; $j++) {
                                        $tempi = $StartDate.AddDays($i).ToString('dd/MM/yyyy')
                                        "D$tempi"
                                        $tempnum = 0 #Avoid creating $0.00 transaction
                                        while ($tempnum -eq 0) {
                                            $tempnum = [math]::Round((Get-Random -Minimum -100000.00 -Maximum 100000.00),2)
                                        }
                                        "T$tempnum"
                                        if ($tempnum -lt 0) { #Determine if transaction is Withdrawal or Deposit
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
            Write-Host $QIFFileName 'is saved to:'(Get-Location).Path
            $response = 'n'
        } until ($response -eq 'n')
        Read-Host -Prompt "Press Enter exit the tool"
        $AllValid = $true
    }
}