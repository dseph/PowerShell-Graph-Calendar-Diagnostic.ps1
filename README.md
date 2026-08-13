# PowerShell-Graph-Calendar-Diagnostic.ps1

Check calendar records for issues within a date range.

`PowerShell-Graph-Calendar-Diagnostic.ps1` reads a mailbox calendar with Microsoft Graph and reports:

* **Meetings with no ending date** – recurring series whose recurrence range is `noEnd`.
* **Meetings over a year long** – a single occurrence longer than `-LongMeetingDays` (365 by default).
* **Over 20 modified meeting occurrences** – recurring series with more than `-MaxModifiedOccurrences`
  exceptions inside the requested range.
* **Duplicate iCalUIDs** – events that share the same `iCalUId`.

## Requirements

* PowerShell 5.1 or PowerShell 7+
* The [Microsoft.Graph.Authentication](https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication) module:

  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```

* The `Calendars.Read` permission for your own mailbox, or `Calendars.Read.All` for another mailbox.

## Usage

```powershell
# Check the signed in user's calendar for the last and next year
.\PowerShell-Graph-Calendar-Diagnostic.ps1

# Check another mailbox for a specific range and write the reports to C:\Reports
.\PowerShell-Graph-Calendar-Diagnostic.ps1 -Mailbox user@contoso.com `
    -StartDate '2024-01-01' -EndDate '2025-01-01' -OutputPath C:\Reports

# Use an existing access token and custom thresholds
.\PowerShell-Graph-Calendar-Diagnostic.ps1 -Mailbox user@contoso.com -AccessToken $token `
    -LongMeetingDays 180 -MaxModifiedOccurrences 10
```

## Output

The following files are created in `-OutputPath`, all prefixed with
`CalendarDiagnostic-<mailbox>-<timestamp>`:

| File | Contents |
| --- | --- |
| `-Summary.txt` | Overall summary with the counts for every check and the list of report files |
| `-NoEndDate.csv` | Meetings with no ending date |
| `-MeetingsOverAYear.csv` | Meetings longer than `-LongMeetingDays` |
| `-ModifiedOccurrences.csv` | Series with more than `-MaxModifiedOccurrences` modified occurrences |
| `-DuplicateICalUIDs.csv` | Events sharing an `iCalUId` |
| `-AllEvents.csv` | Every event that was examined |

## Tests

The [Pester](https://pester.dev) tests in `tests` cover the analysis logic:

```powershell
Invoke-Pester -Path .\tests
```
