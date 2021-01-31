$revision = 'v0.8'
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
# Clear variables if used Globally
$StartDate = ''
$EndDate = ''
$TransactionsPerDay = ''
$TotalDays = ''

# Save Dialog settings
$QIFFileName = 'DummyQIF_' # Default file name to use
$SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveDialog.DefaultExt = 'qif'
$SaveDialog.FileName = $QIFFileName + (Get-Date -Format yyyyMMdd-HHmmss) + '.qif'
$SaveDialog.Filter = 'QIF File (*.qif)|*.qif'

# Summary Values
$TransactionCount = 0

# UI colour formatting | [Forground Label,Background Label,Foreground Val,Background Val]
$DefaultCol = 'gray','black','gray','black'
$ValidCol = 'white',$DefaultCol[1],'green',$DefaultCol[1]
$InvalidCol = 'red',$DefaultCol[1],'red',$DefaultCol[1]
$InFocusCol = 'cyan',$DefaultCol[1],'red',$DefaultCol[1]

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
    try { [DateTime]::ParseExact($String,'d/M/yyyy',[System.Globalization.CultureInfo]'en-NZ')
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
    Write-Host 'QIF-Gen ' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
    Write-Host $revision -NoNewline -ForegroundColor $ValidCol[2] -BackgroundColor $ValidCol[3]
    Write-Host ': Builds a QIF filled with dummy data' -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
    Write-Host
    Write-Host "Start Date: " -NoNewline -ForegroundColor $StartDateCol[0] -BackgroundColor $StartDateCol[1]
    if ($StartDateValid[1] -eq $false) {
        Write-Host $StartDate -ForegroundColor $StartDateCol[2] -BackgroundColor $StartDateCol[3]
    }   else {
        Write-Host $StartDate.ToString("d\/M\/yyyy") -ForegroundColor $StartDateCol[2] -BackgroundColor $StartDateCol[3]
        }
    Write-Host "End Date: " -NoNewline -ForegroundColor $EndDateCol[0] -BackgroundColor $EndDateCol[1]
    if ($EndDateValid[1] -eq $false) {
        Write-Host $EndDate -ForegroundColor $EndDateCol[2] -BackgroundColor $EndDateCol[3]
    }   else {
            Write-Host $EndDate.ToString("d\/M\/yyyy") -ForegroundColor $EndDateCol[2] -BackgroundColor $EndDateCol[3]
            Write-Host 'Days in range: ' -NoNewline -ForegroundColor $EndDateCol[0] -BackgroundColor $EndDateCol[1]
        }
    Write-Host $TotalDays -ForegroundColor $ValidCol[2] -BackgroundColor $ValidCol[3]
    Write-Host
    Write-Host 'Transactions/Day: ' -NoNewline -ForegroundColor $TransactionsCol[0] -BackgroundColor $TransactionsCol[1]
    if ($TransactionsValid -eq $false) {
        Write-Host $TransactionsPerDay -ForegroundColor $TransactionsCol[2] -BackgroundColor $TransactionsCol[3]
    }   else {
            Write-Host $TransactionsPerDay -ForegroundColor $TransactionsCol[2] -BackgroundColor $TransactionsCol[3]
            Write-Host 'Maxmimum possible number of transactions to generate: ' -NoNewline -ForegroundColor $TransactionsCol[0] -BackgroundColor $TransactionsCol[1]
            Write-Host ([int]$TransactionsPerDay * $TotalDays)-ForegroundColor $TransactionsCol[2] -BackgroundColor $TransactionsCol[3]
        }
    Write-Host
}

Update-UI

while ($AllValid -eq $false) {
    if ($StartDateValid[1] -eq $false) {
        $StartDate = Read-Host -Prompt 'Enter start date (d/M/yyyy)' 
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
        $EndDate = Read-Host -Prompt 'Enter end date (d/M/yyyy)'
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
                            $EndDate = $EndDate.ToString("d\/M\/yyyy")
                            Update-UI
                            Write-Host 'End Date' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                            Write-Host ' cannot be before the Start Date' -ForegroundColor $DefaultCol[0] -BackgroundColor $DefaultCol[1]
                        }   else {
                                $EndDate = $EndDateValid[0]
                                $TotalDays = (New-TimeSpan -Start $StartDate -End $EndDate).Days + 1
                                Update-UI
                            }
                    }
            }
    }

    elseif ($TransactionsValid -eq $false) {
        $TransactionsPerDay = Read-Host -Prompt 'Enter maximum random transactions per day'
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
                        Update-UI
                        do {
                            Write-Host 'Create QIF file now [' -NoNewline -ForegroundColor $InFocusCol[0] -BackgroundColor $InFocusCol[1]
                            Write-Host 'Y'-NoNewline -ForegroundColor $ValidCol[2] -BackgroundColor $ValidCol[3] 
                            Write-Host '|' -NoNewline -ForegroundColor $InFocusCol[0] -BackgroundColor $InFocusCol[1]
                            Write-Host 'N' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                            Write-Host ']? ' -NoNewline -ForegroundColor $InFocusCol[0] -BackgroundColor $InFocusCol[1]
                            $response = Read-Host
                            if ($response -eq 'n') {
                                $TransactionsValid = $false
                                Update-UI
                                $response = 'y'
                            }
                        } until ($response -eq 'y')
                    }
            }
    }

    elseif (($StartDateValid -eq $true) -and ($EndDateValid -eq $true) -and ($TransactionsValid -eq $true)) {
        $QIFContent =   for ($i = 0; $i -lt $TotalDays; $i++) {
                            $jcount = 1
                            if ($TransactionsPerDay -gt 1) {
                                $jcount = Get-Random -Minimum 1 -Maximum $TransactionsPerDay
                            }
                            for ($j = 0; $j -lt $jcount; $j++) {
                                $tempi = $StartDate.AddDays($i).ToString('dd\/MM\/yyyy')
                                "D$tempi"
                                $tempnum = 0 #Avoid creating $0.00 transaction
                                while ($tempnum -eq 0) {
                                    $tempnum = [math]::Round((Get-Random -Minimum -10000.00 -Maximum 10000.00),2)
                                }
                                "T$tempnum"
                                if ($tempnum -lt 0) { #Determine if transaction is Withdrawal or Deposit
                                    "PWithdrawal"
                                }   else {
                                        "PDeposit"
                                    }
                                "M"
                                "^"
                                $TransactionCount++
                            }
                        }
        $result = $SaveDialog.ShowDialog()
        if($result -eq 'OK') { 
            '!Type:Bank'| Out-File -FilePath $SaveDialog.FileName -Encoding ascii
            Add-Content $SaveDialog.FileName $QIFContent
            Update-UI
            Write-Host $SaveDialog.FileName -NoNewline -ForegroundColor $InFocusCol[0] -BackgroundColor $InFocusCol[1]
            Write-Host ' has been created with ' -NoNewline -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
            Write-Host $TransactionCount -NoNewline -ForegroundColor $ValidCol[2] -BackgroundColor $ValidCol[3]
            Write-Host ' transactions.' -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
            Write-Host
            Read-Host -Prompt "Press Enter exit the tool"
            $AllValid = $true
        }   else {
                Write-Host 'File creation cancelled' -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                do {
                    Write-Host '[' -NoNewline -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
                    Write-Host 'R' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                    Write-Host ']estart or [' -NoNewline -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
                    Write-Host 'Q' -NoNewline -ForegroundColor $InvalidCol[0] -BackgroundColor $InvalidCol[1]
                    Write-Host ']uit? ' -NoNewline -ForegroundColor $ValidCol[0] -BackgroundColor $ValidCol[1]
                    $response = Read-Host
                    if ($response -eq 'q') {
                        $AllValid = $true
                        $response = 'r'
                    }   elseif ($response -eq 'r') {
                            $StartDate = ''
                            $EndDate = ''
                            $TransactionsPerDay = ''
                            $TotalDays = ''
                            $StartDateValid = $null,$false
                            $EndDateValid = $null,$false
                            $TransactionsValid = $false
                            $AllValid = $false
                            Update-UI
                        }
                } until ($response -eq 'r')
            }
    }
}