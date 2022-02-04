# HsePowerShell220202

## Useful Links
- Create a lab environment with PowerShell: [AutomatedLab](https://automatedlab.org/en/latest/)
- [Regex101 builder](https://regex101.com/). If you need to parse complex text, RegEx is your friend.
- [From a one-liner to a full-featured PowerShell advanced function](https://github.com/raandree/PowerShellTraining). Recommended if you want to learn all the features PowerShell provide to organize code in function.
- [PowerShell Explained with Kevin Marquette
](https://powershellexplained.com/sitemap/?utm_source=blog&utm_medium=blog&utm_content=recent)

&nbsp;

## PowerShell Cheat Sheets
  - https://cdn.comparitech.com/wp-content/uploads/2018/08/Comparitech-Powershell-cheatsheet.pdf
  - https://ramblingcookiemonster.github.io/images/Cheat-Sheets/powershell-basic-cheat-sheet2.pdf
  - https://gitlab.com/JamesHedges/notes/-/wikis/Powershell/PowerShell-Cheat-Sheet
  - https://www.theochem.ru.nl/~pwormer/teachmat/PS_cheat_sheet.html

&nbsp;

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
&nbsp;

- ### Simpler Error Handling

    Get all files and folders, don't show the errors in the console but save them in the variable `$myError`.

    ```powershell
    $data = dir -Recurse -ErrorAction SilentlyContinue -ErrorVariable myError
    $data.Count #>277757
    $myError.Count #>46

    #This line returns all files and folders that could not be read
    $myError.TargetObject
    ```

&nbsp;

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

&nbsp;

- ### Find all files that have been changed in a certain time frame

    ```powershell
    $start = Get-Date -Date '2.2.2022 08:00:00'
    $end = Get-Date -Date '2.2.2022 13:00:00'

    dir d:\ -Recurse | Where-Object { $_.LastWriteTime -gt $start -and $_.LastWriteTime -lt $end }
    ```

&nbsp;

- ### Get all text files that contain the string `a877777`.
    ```powershell
    dir -Path E:\LabSources -Filter *.txt -Recurse |
        Where-Object {
            Get-Content -Path $_.FullName | Where-Object { $_ -like '*a877777*'
        }
    }
    ```

&nbsp;

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

    &nbsp;

   - ### `Where-Object` and `Group-Object` with Active Directory
    ```powershell
    #Get all users working in Bolivia by filtering them

    #Get one object first to get an idea how it looks like and where the information about
    #the country is actually stored
    Get-ADUser -Identity a174978 -Properties *

    #you will see it is in the property 'c' or 'country' and not like you may think in the
    #property 'country/region' as displayed in dsa.msc. You also realize that the country
    #AD is stored as the country code (BO) and not the country's display name (Bolivia)

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

&nbsp;

- ### Import and Export CSV and XML by using processes

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

&nbsp;

- ### Get the PSReadline History in Notepad
    ```powershell
    notepad (Get-PSReadLineOption).HistorySavePath
    ```

&nbsp;

- ### Use `Out-GridView` to allow someone to select processes that will be stopped
    ```powershell
    Get-Process | Out-GridView                          
    Get-Process | Out-GridView -PassThru                
    Get-Process | Out-GridView -PassThru | Stop-Process 
    ```

&nbsp;

- ### Get Logon Events from the Security Event Log

    Get all logon events for the user `someone` that have been generated after 18:00 today

    ```powershell
    Get-EventLog -LogName Security -InstanceId 4624 | Where-Object { $_.ReplacementStrings[5] -eq 'someone' -and $_.TimeGenerated -gt (Get-Date -Hour 18 -Minute 0 -Second 0).AddDays(-1) }
    ```

    Same thing but this command uses dynamic columns to represent the data in a better way.

    This time we are getting the events from yesterday (`.AddDays(-1)`).

    ```powershell
    Get-EventLog -LogName Security -InstanceId 4624 |
    Where-Object {
    $_.TimeGenerated -gt (Get-Date -Hour 18 -Minute 0 -Second 0).AddDays(-1) } | 
    Format-Table -Property TimeGenerated,
    @{ Name = 'Username'; Expression = { $_.ReplacementStrings[5] } }, 
    @{ Name = 'AuthType'; Expression = { $_.ReplacementStrings[10] } }, 
    @{ Name = 'Domain'; Expression = { $_.ReplacementStrings[6] } }
    ```

&nbsp;

- ### Check the status of a service and start it when stopped.

    ```powershell
    $serviceName = 'W32Time'
    $s = Get-Service -Name $serviceName
    if ($s.Status -ne 'Running') {
        Write-Host "Service '$serviceName' is not running, starting it..." -NoNewline
        Start-Service -Name $serviceName
        Write-Host 'done'
    }
    else {
        Write-Host "Service '$serviceName' is already running"
    }
    ```

- Look for a certain pattern in a text file and print out all lines matching that pattern including the line number. Also each match gets a context added so you can access the two previous and following lines.

    ```powershell
    $content = Get-Content -Path C:\Windows\debug\WIA\wiatrace.log | Select-String -Pattern 'An error occured' -Context 2, 2
    $content | Format-Table -Property LineNumber, @{ Name = 'LineContent'; Expression = { $_ } }
    $content[6].Context.PreContext
    $content[6].Context.PostContext
    ```

&nbsp;

- ### ForEach samples 
    - Start all VMs that are not already running yet.

        ```powershell
        $vms = Get-VM -Name DSC*

        $vms | ForEach-Object {
            if ($_.State -ne 'Running') {
                Write-Host "Starting VM $($_.Name)..." -NoNewline
                $_ | Start-VM
                Write-Host 'done'
            }
            else {
                Write-Host "VM $($_.Name) is already started"
            }
        }
        ```

        &nbsp;

    - Create 100 test users in Active Directory

        > Note: The `f` operator allows some special formatting, in this case the expanding of the integer to 3 digits.

        ```powershell
        1..100 | ForEach-Object {
            $name = "TestUser {0:D3}" -f $_
            New-ADUser -Name $name -Path 'OU=Demo2,DC=contoso,DC=com'
        }
        ```

        &nbsp;

    - You don't necessarily need a ForEach loop in PowerShell to process a number of elements as the pipe (`|`) is implements a ForEach loop.

        ```powershell
        $users = Get-ADUser -Filter * -SearchBase 'OU=Demo2,DC=contoso,DC=com'
        $users | Set-ADUser -Description 'Test User'
        ```

    &nbsp;

    - You can use the `ForEach-Object` cmdlet or the `foreach` keyword to create data in Active Directory but many cmdlets accept pipeline input by property name (ValueFromPipelineByPropertyName), hence a foreach loop is not required

        ```powershell
        Import-Module ActiveDirectory
        
        $Domain="@abc.com"
        $UserOu="OU=Users,DC=abc,DC=com"
        $NewUsersList=Import-CSV "D:\userstobeimported.csv"
        
        ForEach ($User in $NewUsersList) 
        {
            $givenName=$User.givenName
            $sAMAccountName=$User.sAMAccountName
            $userPrincipalName=$User.sAMAccountName+$Domain
            $userPassword=$User.Password
            $expire=$null

            New-ADUser -Name $givenName -GivenName $givenName  -SamAccountName $sAMAccountName 
        }
        ```

        Using the PowerShell pipeline feature `ValueFromPipelineByPropertyName` shortens the script to just one line (two lines if you want to create an OU as well).

        ```powershell
        $ou = New-ADOrganizationalUnit -Name Demo1 -PassThru
        Import-Csv -Path .\People.csv -Delimiter ';' | New-ADUser -Path $ou
        ```

        Removing the users works like this:

        ```powershell
        Get-ADUser -Filter 'Description -like "Class*"' | Remove-ADUser -Confirm:$false
        ```

    &nbsp;

    - ### Use the `split` and `join` operator to bring list into a new format

        ```powershell
        $emails = 'Hall.Nichols@contoso.com, Hoyt.Cortez@contoso.com, Adrienne.Bean@contoso.com, Oliver.Cote@contoso.com, Sasha.Glenn@contoso.com, Wing.Foreman@contoso.com, Phillip.Joyner@contoso.com, Miranda.Espinoza@contoso.com, Brody.Shelton@contoso.com, Reece.Joseph@contoso.com'
        $emails = $emails -split ',' | ForEach-Object { $_.Trim() }
        '"' + ($emails -join '"; "') + '"'
        ```

- ### Arrays are immutable types meaning they cannot be extended in length. The `+=` operator makes it look like it is possible but PowerShell copies the growing data as often in memory as you add something to the list. This can take hours. A very fast alternative is using an `System.Collections.ArrayList`.
    &nbsp;

    ```powershell
    $files = @()

    1..100000 | ForEach-Object {
        $files += "New File $_"
    }

    $files.Count

    #----------------------------------------

    $al = New-Object System.Collections.ArrayList

    1..100000 | ForEach-Object {
        $null = $al.Add("New File $_") #| Out-Null
    }

    $al.Count
    ```

&nbsp;

- ### Hashtables: Store data about computers in a hashtable, then export and import it via JSON.
    &nbsp;

    ```powershell
    $memory1 = 2gb
    $disk1 = '1TB SSD'
    $cpu1 = 'i7'
    $name1 = 'My Computer'
    $assetTag1 = 12334

    $memory2 = 2gb
    $disk2 = '1TB SSD'
    $cpu2 = 'i7'
    $name2 = 'My Computer'
    $assetTag2 = 12334

    $computers = @{
        Name = 'Office Machine 1'
        Memory = 8GB
        CPU = 'Intel i7'
        AssetTag = 'x123'
        Disks = @{
            Type = 'SSD'
            Size = 1TB
        },
        @{
            Type = 'HDD'
            Size = 12TB
        }
    },
    @{
        Name = 'Office Machine 2'
        Memory = 16GB
        CPU = 'Intel i97'
        AssetTag = 'x123'
        Disks = @{
            Type = 'SSD'
            Size = 41TB
        },
        @{
            Type = 'HDD'
            Size = 12TB
        },
        @{
            Type = 'HDD'
            Size = 12TB
        }
    }

    #get the size of the first disk of the first computer
    $computers[0].Disks[0].Size

    #remove disk 3 from computer 1
    $computers[1].Disks = $computers[1].Disks[0, 1]
    $computers[1].Disks.Count

    $computers | ConvertTo-Json | Out-File -FilePath D:\Computers.json

    $computers = Get-Content -Path D:\Computers.json -Raw | ConvertFrom-Json
    ```

&nbsp;

- ### Find duplicate files by comparing the hash created by the file content using `Get-FileHash`

    ```powershell
    dir -Recurse | Group-Object -Property { ($_ | Get-FileHash).Hash } | Where-Object Count -gt 1
    ```

&nbsp;

- ### Remoting: Uninstall and install Notepad++ on all machines in a certain OU.

    Some more details about remote software installation with PowerShell:
    - [Working with Software Installations](https://docs.microsoft.com/en-us/powershell/scripting/samples/working-with-software-installations?view=powershell-7.2)
    - [**Powershell: Remote install software**](https://powershellexplained.com/2017-04-22-Powershell-installing-remote-software/)

    &nbsp;

    ```powershell
        $command = {
        if (Test-Path -Path 'C:\Program Files\Notepad++\notepad++.exe')
        {
            Write-Host 'Notpad+ is installed and will be uninstalled...' -NoNewline
            $p = Start-Process -FilePath 'C:\Program Files\Notepad++\uninstall.exe' -ArgumentList /S -PassThru
            $p.WaitForExit()
            Write-Host done
        }
        else
        {
            Write-Host 'Notpad+ is not installed'
        }

        Write-Host 'Installing Notepad++...' -NoNewline
        $p = Start-Process -FilePath C:\npp.8.1.9.Installer.x64.exe -ArgumentList /S -PassThru
        $p.WaitForExit()
        Write-Host done
    }
    $computers = Get-ADComputer -Filter * -SearchBase 'CN=Computers,DC=contoso,DC=com'

    Invoke-Command -ComputerName $computers.DnsHostName -ScriptBlock $command
    ```

&nbsp;

- ### Different kind of loops

    ```powershell
    $a = 1..5

    $a | ForEach-Object {
        "Test $_"
    }

    foreach ($value in $a) {
        "Test $value"
    }

    for ($i = 0; $i -lt $a.Length; $i++) {
        "Test $($a[$i])"
    }

    Write-Host "Waiting for file 'D:\Untitled6.ps1' to be present'" -NoNewline
    do {
        Start-Sleep -Milliseconds 500
        Write-Host . -NoNewline
    } until (Test-Path D:\Untitled6.ps1)
    Write-Host "File found"
    ```
