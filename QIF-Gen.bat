#RUNS THE LATEST VERSION FROM GITHUB
powershell -Command Invoke-Expression (Invoke-WebRequest https://raw.githubusercontent.com/themalse/QIF-Gen/master/QIF-Gen.ps1).Content
