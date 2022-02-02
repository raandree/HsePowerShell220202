# HsePowerShell220202

## Useful Links
- Create a lab environment with PowerShell: [AutomatedLab](https://automatedlab.org/en/latest/)
- [Regex101 builder](https://regex101.com/). If you need to parse complex text, RegEx is your friend.

## PowerShell Cheat Sheets
  - https://cdn.comparitech.com/wp-content/uploads/2018/08/Comparitech-Powershell-cheatsheet.pdf
  - https://ramblingcookiemonster.github.io/images/Cheat-Sheets/powershell-basic-cheat-sheet2.pdf
  - https://gitlab.com/JamesHedges/notes/-/wikis/Powershell/PowerShell-Cheat-Sheet
  - https://www.theochem.ru.nl/~pwormer/teachmat/PS_cheat_sheet.html


## Code Snippets

- Get all outlook events from the application event log (local machine)

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

- Simpler Error Handling

    Get all files and folders, don't show the errors in the console but save them in the variable `$myError`.

    ```powershell
    $data = dir -Recurse -ErrorAction SilentlyContinue -ErrorVariable myError
    $data.Count #>277757
    $myError.Count #>46

    #This line returns all files and folders that could not be read
    $myError.TargetObject
    ```

- DateTime Object
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
