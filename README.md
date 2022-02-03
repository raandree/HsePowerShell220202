# HsePowerShell220202

## Useful Links
- Create a lab environment with PowerShell: [AutomatedLab](https://automatedlab.org/en/latest/)
- [Regex101 builder](https://regex101.com/). If you need to parse complex text, RegEx is your friend.
- [From a one-liner to a full-featured PowerShell advanced function](https://github.com/raandree/PowerShellTraining). Recommended if you want to learn all the features PowerShell provide to organize code in function.

## PowerShell Cheat Sheets
  - https://cdn.comparitech.com/wp-content/uploads/2018/08/Comparitech-Powershell-cheatsheet.pdf
  - https://ramblingcookiemonster.github.io/images/Cheat-Sheets/powershell-basic-cheat-sheet2.pdf
  - https://gitlab.com/JamesHedges/notes/-/wikis/Powershell/PowerShell-Cheat-Sheet
  - https://www.theochem.ru.nl/~pwormer/teachmat/PS_cheat_sheet.html


## Code Snippets

- ### Get all outlook events from the application event log (local machine)

    ```powershell
    Get-EventLog -LogName Application -EntryType Error, Warning -Source Outlook
    ```

    If you want to export the result, `Export-Csv` is probably the best way

    ```powershell
    Get-EventLog -LogName Application -EntryType Error, Warning -Source Outlook | Export-Csv -Path .\Events.csv
    ```

    If you want these events only from the last two hours

    ```powershell
    Get-EventLog -LogName Application -EntryType Error, Warning -Source Outlook -After (Get-Date).AddHours(-2)
    ```

- ### Simpler Error Handling

    Get all files and folders, don't show the errors in the console but save them in the variable `$myError`.

    ```powershell
    $data = dir -Recurse -ErrorAction SilentlyContinue -ErrorVariable myError
    $data.Count #>277757
    $myError.Count #>46

    #This line returns all files and folders that could not be read
    $myError.TargetObject
    ```

- ### DateTime Object
    ```powershell
    $d = Get-Date
    $d.Ticks    
    $d.AddTicks($d.Ticks)
    #Jump back in time where .net / PowerShell starts counting the ticks.
    $d.AddTicks(-$d.Ticks)

    $d.AddHours(250)

    #Calculate the remaining time in this year.
    $nextYear = Get-Date -Date '1. 1. 2023'
    $nextYear - $d
    ```

- ### Find all files that have been changed in a certain time frame

    ```powershell
    $start = Get-Date -Date '2.2.2022 08:00:00'
    $end = Get-Date -Date '2.2.2022 13:00:00'

    dir d:\ -Recurse | Where-Object { $_.LastWriteTime -gt $start -and $_.LastWriteTime -lt $end }
    ```

- ### Get all text files that contain the string `a877777`.
    ```powershell
    dir -Path E:\LabSources -Filter *.txt -Recurse |
        Where-Object {
            Get-Content -Path $_.FullName | Where-Object { $_ -like '*a877777*'
        }
    }
    ```

- ### Group all files by the year of their creation
    ```powershell
    dir -Recurse | Group-Object -Property { $_.CreationTime.ToString('yyyy') }
    ```
    ```
    Count Name                      Group                                                                        
    ----- ----                      -----                                                                        
    1244 2021                      {CustomRoles, DscConfigurations, GitRepositories, ISOs...}                   
    170 2022                      {HostStart.ps1, AssemblyInfo.cs, ExchangeServer2016-x64-CU20.ISO, Exchange...
    44 2020                      {en_sql_server_2012_standard_edition_with_service_pack_4_x64_dvd_100075944...
   ```

   - ### `Where-Object` and `Group-Object` with Active Directory
    ```powershell
    #Get all users working in Bolivia by filtering them

    #Get one object first to get an idea how it looks like and where the information about
    #the country is actually stored
    Get-ADUser -Identity a174978 -Properties *

    #you will see it is in the property 'c' or 'country' and not like you may think in the
    #property 'country/region' as displayed in dsa.msc. You also realize that the country
    #AD is stored as the country code (BO) and not the countrie's display name (Bolivia)

    $users = Get-ADUser -Filter * -Properties Country | Where-Object Country -eq BO
    $users.Count

    #Filtering should be always done - if possible - on the system that creates the data
    #snapshot / stores the data.
    $users = Get-ADUser -Filter { Country -eq "BO" }
    $users.Count

    #It could be also interesting to get all users and group them by the country. Then you
    #get not only the employee count of Bolivia but every country that is represented in AD.
    $users = Get-ADUser -Filter * -Properties Country | Group-Object -Property Country | Sort-Object -Property Name
    ```

- Import and Export CSV and XML by using processes

    ```powershell
    Get-Process | Select-Object -Property Name, ID, Threads, StartInfo | Export-Csv -Path p.csv -NoTypeInformation
    start .\p.csv
    $p = Import-Csv -Path D:\p.csv
    $p[0].Threads #returns just 'System.Diagnostics.ProcessThreadCollection'

    Get-Process | Select-Object -Property Name, ID, Threads, StartInfo | Export-Clixml -Path p.xml
    dir #the XML file is much larger then CSV
    $p = Import-Clixml D:\p.xml
    #this time we can access all data and the other dimensions are preserved.
    $p[0]
    $p[0].Threads.Count
    $p[0].Threads
    $p[0].StartInfo
    ```

- Get the PSReadline History in Notepad
    ```powershell
    notepad (Get-PSReadLineOption).HistorySavePath
    ```

- Use `Out-GridView` to allow someone to select processes that will be stopped
    ```powershell
    Get-Process | Out-GridView                          
    Get-Process | Out-GridView -PassThru                
    Get-Process | Out-GridView -PassThru | Stop-Process 
    ```

- Get Logon Events from the Security Event Log

    Get all logon events for the user `someone` that have been generated after 18:00 today

    ```powershell
    Get-EventLog -LogName Security -InstanceId 4624 | Where-Object { $_.ReplacementStrings[5] -eq 'someone' -and $_.TimeGenerated -gt (Get-Date -Hour 18 -Minute 0 -Second 0).AddDays(-1) }
    ```

    Same thing but this command uses dynamic columns to represent the data in a better way.

    This time we are getting the events from yesterday (`.AddDays(-1)`).

    ```powershell
    Get-EventLog -LogName Security -InstanceId 4624 | Where-Object { $_.TimeGenerated -gt (Get-Date -Hour 18 -Minute 0 -Second 0).AddDays(-1) } | Format-Table -Property TimeGenerated, @{ Name = 'Username'; Expression = { $_.ReplacementStrings[5] } }, @{ Name = 'AuthType'; Expression = { $_.ReplacementStrings[10] } }, @{ Name = 'Domain'; Expression = { $_.ReplacementStrings[6] } }
    ```
