<#
.SYNOPSIS
    Checks a mailbox calendar for common problem records by using Microsoft Graph.

.DESCRIPTION
    Reads the calendar of the specified mailbox and reports the following issues:

        * Meetings with no ending date (recurring series with a "noEnd" recurrence range).
        * Meetings that are longer than a year (a single occurrence spanning more than
          the number of days specified by -LongMeetingDays).
        * Recurring series with more than -MaxModifiedOccurrences modified occurrences
          (exceptions) inside the requested date range.
        * Events that share the same iCalUId (duplicates).

    An overall summary file is created together with one .csv file per issue type.

.PARAMETER Mailbox
    User principal name (or id) of the mailbox to check. Defaults to the signed in user.

.PARAMETER StartDate
    Beginning of the range to check. Defaults to one year before today.

.PARAMETER EndDate
    End of the range to check. Defaults to one year after today.

.PARAMETER OutputPath
    Folder where the report files are written. Defaults to the current directory.

.PARAMETER LongMeetingDays
    A single meeting longer than this number of days is reported. Defaults to 365.

.PARAMETER MaxModifiedOccurrences
    A recurring series with more than this number of modified occurrences is reported.
    Defaults to 20.

.PARAMETER AccessToken
    Optional Microsoft Graph access token. When it is not supplied the script uses the
    current Connect-MgGraph session, and connects if there is none.

.EXAMPLE
    .\PowerShell-Graph-Calendar-Diagnostic.ps1 -Mailbox user@contoso.com

.EXAMPLE
    .\PowerShell-Graph-Calendar-Diagnostic.ps1 -Mailbox user@contoso.com -StartDate '2024-01-01' -EndDate '2025-01-01' -OutputPath C:\Reports

.NOTES
    Requires the Microsoft.Graph.Authentication module and the Calendars.Read
    (or Calendars.Read.All) permission.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $Mailbox,

    [datetime] $StartDate = (Get-Date).Date.AddYears(-1),

    [datetime] $EndDate = (Get-Date).Date.AddYears(1),

    [string] $OutputPath = (Get-Location).Path,

    [ValidateRange(1, 36500)]
    [int] $LongMeetingDays = 365,

    [ValidateRange(0, 10000)]
    [int] $MaxModifiedOccurrences = 20,

    [string] $AccessToken
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers

function ConvertTo-GraphDateTime {
    <#
    .SYNOPSIS
        Converts a Graph dateTimeTimeZone value (or a string) into a [datetime].
    #>
    [CmdletBinding()]
    [OutputType([System.Nullable[datetime]])]
    param(
        [Parameter(Position = 0)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return $null }

    if ($Value -is [datetime]) { return $Value }

    $text = $null
    if ($Value -is [string]) {
        $text = $Value
    }
    else {
        $dateTime = Get-EventProperty -CalendarEvent $Value -Name 'dateTime'
        if ($null -eq $dateTime) { $dateTime = Get-EventProperty -CalendarEvent $Value -Name 'DateTime' }
        if ($null -ne $dateTime) { $text = [string]$dateTime }
    }

    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor `
              [System.Globalization.DateTimeStyles]::AdjustToUniversal

    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref] $parsed)) {
        return $parsed
    }

    Write-Warning "Unable to parse date/time value '$text'."
    return $null
}

function Get-EventProperty {
    <#
    .SYNOPSIS
        Safely reads a (possibly missing) property from a Graph event object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent,

        [Parameter(Mandatory, Position = 1)]
        [string] $Name
    )

    if ($null -eq $CalendarEvent) { return $null }

    if ($CalendarEvent -is [System.Collections.IDictionary]) {
        if ($CalendarEvent.Contains($Name)) { return $CalendarEvent[$Name] }
        return $null
    }

    $property = $CalendarEvent.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }

    return $property.Value
}

function Get-EventRecurrenceRange {
    <#
    .SYNOPSIS
        Returns the recurrence range of an event, or $null when it is not recurring.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent
    )

    $recurrence = Get-EventProperty -CalendarEvent $CalendarEvent -Name 'recurrence'
    if ($null -eq $recurrence) { return $null }

    return Get-EventProperty -CalendarEvent $recurrence -Name 'range'
}

function Test-EventNoEndDate {
    <#
    .SYNOPSIS
        Returns $true when the event is a recurring series without an end date.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent
    )

    $range = Get-EventRecurrenceRange -CalendarEvent $CalendarEvent
    if ($null -eq $range) { return $false }

    $type = [string](Get-EventProperty -CalendarEvent $range -Name 'type')

    return ($type -eq 'noEnd')
}

function Get-EventDuration {
    <#
    .SYNOPSIS
        Returns the length, in days, of a single occurrence of the event.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent
    )

    $start = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'start')
    $end = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'end')

    if ($null -eq $start -or $null -eq $end) { return 0 }

    return ($end - $start).TotalDays
}

function Test-EventTooLong {
    <#
    .SYNOPSIS
        Returns $true when a single occurrence of the event is longer than $LongMeetingDays.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent,

        [int] $LongMeetingDays = 365
    )

    return ((Get-EventDuration -CalendarEvent $CalendarEvent) -gt $LongMeetingDays)
}

function Test-EventInRange {
    <#
    .SYNOPSIS
        Returns $true when the event (or its recurrence range) overlaps the requested range.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent,

        [Parameter(Mandatory)]
        [datetime] $StartDate,

        [Parameter(Mandatory)]
        [datetime] $EndDate
    )

    $start = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'start')
    $end = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'end')

    $range = Get-EventRecurrenceRange -CalendarEvent $CalendarEvent
    if ($null -ne $range) {
        $rangeStart = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $range -Name 'startDate')
        $rangeEnd = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $range -Name 'endDate')
        $rangeType = [string](Get-EventProperty -CalendarEvent $range -Name 'type')

        if ($null -ne $rangeStart) { $start = $rangeStart }

        if ($rangeType -eq 'noEnd' -or $rangeType -eq 'numbered') {
            # The last occurrence is unknown without expanding the series, assume it is open ended.
            $end = [datetime]::MaxValue
        }
        elseif ($null -ne $rangeEnd) {
            $end = $rangeEnd
        }
    }

    if ($null -eq $start -and $null -eq $end) { return $true }
    if ($null -eq $start) { $start = $end }
    if ($null -eq $end) { $end = $start }

    return ($start -le $EndDate -and $end -ge $StartDate)
}

function Get-DuplicateICalUid {
    <#
    .SYNOPSIS
        Returns the groups of events that share the same iCalUId.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [AllowNull()]
        [object[]] $Events
    )

    if ($null -eq $Events) { return @() }

    $withUid = @($Events | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string](Get-EventProperty -CalendarEvent $_ -Name 'iCalUId'))
        })

    if ($withUid.Count -eq 0) { return @() }

    return @($withUid |
        Group-Object -Property { [string](Get-EventProperty -CalendarEvent $_ -Name 'iCalUId') } |
        Where-Object { $_.Count -gt 1 })
}

function ConvertTo-DiagnosticRecord {
    <#
    .SYNOPSIS
        Flattens a Graph event into the object written to the .csv reports.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent,

        [int] $ModifiedOccurrenceCount = 0
    )

    $range = Get-EventRecurrenceRange -CalendarEvent $CalendarEvent
    $rangeType = if ($null -eq $range) { '' } else { [string](Get-EventProperty -CalendarEvent $range -Name 'type') }
    $rangeEnd = if ($null -eq $range) { $null } else { Get-EventProperty -CalendarEvent $range -Name 'endDate' }
    $organizer = Get-EventProperty -CalendarEvent $CalendarEvent -Name 'organizer'
    $organizerAddress = ''
    if ($null -ne $organizer) {
        $emailAddress = Get-EventProperty -CalendarEvent $organizer -Name 'emailAddress'
        if ($null -ne $emailAddress) {
            $organizerAddress = [string](Get-EventProperty -CalendarEvent $emailAddress -Name 'address')
        }
    }

    $start = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'start')
    $end = ConvertTo-GraphDateTime (Get-EventProperty -CalendarEvent $CalendarEvent -Name 'end')

    return [pscustomobject] @{
        Subject                 = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'subject')
        Organizer               = $organizerAddress
        StartUtc                = $start
        EndUtc                  = $end
        DurationDays            = [math]::Round((Get-EventDuration -CalendarEvent $CalendarEvent), 2)
        EventType               = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'type')
        RecurrenceRangeType     = $rangeType
        RecurrenceEndDate       = [string]$rangeEnd
        ModifiedOccurrenceCount = $ModifiedOccurrenceCount
        ICalUId                 = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'iCalUId')
        WebLink                 = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'webLink')
        Id                      = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'id')
    }
}

#endregion Helpers

#region Microsoft Graph access

function Connect-DiagnosticGraph {
    <#
    .SYNOPSIS
        Makes sure that a Microsoft Graph session is available.
    #>
    [CmdletBinding()]
    param(
        [string] $AccessToken
    )

    if (-not [string]::IsNullOrWhiteSpace($AccessToken)) { return }

    if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
        throw 'The Microsoft.Graph.Authentication module is required. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $context = $null
    try { $context = Get-MgContext } catch { $context = $null }

    if ($null -eq $context) {
        Write-Verbose 'Connecting to Microsoft Graph.'
        Connect-MgGraph -Scopes 'Calendars.Read' -NoWelcome -ErrorAction Stop | Out-Null
    }
}

function Invoke-DiagnosticGraphRequest {
    <#
    .SYNOPSIS
        Performs a GET against Microsoft Graph and follows @odata.nextLink.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string] $Uri,

        [string] $AccessToken
    )

    $headers = @{ 'Prefer' = 'outlook.timezone="UTC"' }
    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        Write-Verbose "GET $next"

        if ([string]::IsNullOrWhiteSpace($AccessToken)) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $headers -OutputType PSObject
        }
        else {
            $requestHeaders = $headers.Clone()
            $requestHeaders['Authorization'] = 'Bearer ' + $AccessToken
            $response = Invoke-RestMethod -Method GET -Uri $next -Headers $requestHeaders
        }

        $value = Get-EventProperty -CalendarEvent $response -Name 'value'
        if ($null -ne $value) {
            foreach ($item in @($value)) { $results.Add($item) }
        }
        elseif ($null -ne $response) {
            $results.Add($response)
        }

        $next = [string](Get-EventProperty -CalendarEvent $response -Name '@odata.nextLink')
    }

    return $results.ToArray()
}

function Get-MailboxUriPrefix {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string] $Mailbox
    )

    if ([string]::IsNullOrWhiteSpace($Mailbox)) { return 'https://graph.microsoft.com/v1.0/me' }

    return "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Mailbox))"
}

function Get-CalendarEvent {
    <#
    .SYNOPSIS
        Returns the calendar events (single instances and series masters) of a mailbox.
    #>
    [CmdletBinding()]
    param(
        [string] $Mailbox,

        [string] $AccessToken
    )

    $select = 'id,iCalUId,subject,organizer,start,end,type,recurrence,webLink,isCancelled'
    $uri = '{0}/events?$select={1}&$top=100' -f (Get-MailboxUriPrefix -Mailbox $Mailbox), $select

    return Invoke-DiagnosticGraphRequest -Uri $uri -AccessToken $AccessToken
}

function Get-ModifiedOccurrenceCount {
    <#
    .SYNOPSIS
        Returns the number of modified occurrences (exceptions) of a recurring series.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory, Position = 0)]
        $CalendarEvent,

        [Parameter(Mandatory)]
        [datetime] $StartDate,

        [Parameter(Mandatory)]
        [datetime] $EndDate,

        [string] $Mailbox,

        [string] $AccessToken
    )

    $id = [string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'id')
    if ([string]::IsNullOrWhiteSpace($id)) { return 0 }

    $uri = '{0}/events/{1}/instances?startDateTime={2}&endDateTime={3}&$select=id,type,start,end&$top=100' -f `
        (Get-MailboxUriPrefix -Mailbox $Mailbox),
        [uri]::EscapeDataString($id),
        $StartDate.ToString('s'),
        $EndDate.ToString('s')

    try {
        $instances = Invoke-DiagnosticGraphRequest -Uri $uri -AccessToken $AccessToken
    }
    catch {
        Write-Warning "Unable to read the occurrences of '$([string](Get-EventProperty -CalendarEvent $CalendarEvent -Name 'subject'))': $($_.Exception.Message)"
        return 0
    }

    return @($instances | Where-Object {
            [string](Get-EventProperty -CalendarEvent $_ -Name 'type') -eq 'exception'
        }).Count
}

#endregion Microsoft Graph access

#region Reporting

function Get-DiagnosticRecordHeader {
    <#
    .SYNOPSIS
        Returns the .csv header line used by the reports.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $template = ConvertTo-DiagnosticRecord -CalendarEvent ([pscustomobject] @{})

    return (($template.PSObject.Properties.Name | ForEach-Object { '"{0}"' -f $_ }) -join ',')
}

function Format-DiagnosticCount {
    <#
    .SYNOPSIS
        Returns the aligned "check : count" lines of the summary.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [hashtable] $Result
    )

    $counters = [ordered] @{
        'Events examined'                                                              = $Result.TotalEvents
        'Meetings with no ending date'                                                 = @($Result.NoEndDate).Count
        "Meetings longer than $($Result.LongMeetingDays) days"                         = @($Result.LongMeetings).Count
        "Series with more than $($Result.MaxModifiedOccurrences) modified occurrences" = @($Result.ModifiedOccurrences).Count
        'Events with a duplicate iCalUId'                                              = @($Result.DuplicateICalUIDs).Count
    }

    $width = ($counters.Keys | Measure-Object -Property Length -Maximum).Maximum

    return @(foreach ($key in $counters.Keys) { '{0} : {1}' -f $key.PadRight($width), $counters[$key] })
}

function Write-DiagnosticReport {
    <#
    .SYNOPSIS
        Writes the summary file and one .csv file per issue type.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [hashtable] $Result,

        [Parameter(Mandatory)]
        [string] $OutputPath
    )

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $mailboxTag = $Result.Mailbox
    if ([string]::IsNullOrWhiteSpace($mailboxTag)) { $mailboxTag = 'me' }
    $mailboxTag = ($mailboxTag -replace '[^\w\.\-]', '_')
    $prefix = Join-Path -Path $OutputPath -ChildPath ("CalendarDiagnostic-{0}-{1}" -f $mailboxTag, $stamp)

    $files = [ordered] @{
        NoEndDate           = "$prefix-NoEndDate.csv"
        LongMeetings        = "$prefix-MeetingsOverAYear.csv"
        ModifiedOccurrences = "$prefix-ModifiedOccurrences.csv"
        DuplicateICalUIDs   = "$prefix-DuplicateICalUIDs.csv"
        AllEvents           = "$prefix-AllEvents.csv"
    }

    $written = [System.Collections.Generic.List[string]]::new()

    foreach ($name in $files.Keys) {
        $records = @($Result[$name])
        $path = $files[$name]

        if ($PSCmdlet.ShouldProcess($path, 'Write report')) {
            if ($records.Count -eq 0) {
                # Keep a header only .csv so that every check has a matching, importable report.
                Set-Content -LiteralPath $path -Value (Get-DiagnosticRecordHeader) -Encoding UTF8
            }
            else {
                $records | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
            }
        }

        $written.Add($path)
    }

    $summaryPath = "$prefix-Summary.txt"

    $counts = Format-DiagnosticCount -Result $Result
    $summary = @(
        'Calendar Diagnostic Summary'
        '==========================='
        "Generated (UTC)          : $((Get-Date).ToUniversalTime().ToString('u'))"
        "Mailbox                  : $($Result.Mailbox)"
        "Range start (UTC)        : $($Result.StartDate.ToString('u'))"
        "Range end (UTC)          : $($Result.EndDate.ToString('u'))"
        ''
        $counts
        ''
        'Report files'
        '------------'
    ) + @($written) + @($summaryPath)

    if ($PSCmdlet.ShouldProcess($summaryPath, 'Write summary')) {
        Set-Content -LiteralPath $summaryPath -Value $summary -Encoding UTF8
    }

    $written.Add($summaryPath)

    return $written.ToArray()
}

#endregion Reporting

function Invoke-CalendarDiagnostic {
    <#
    .SYNOPSIS
        Analyses the supplied (or downloaded) events and returns the diagnostic result.
    #>
    [CmdletBinding()]
    param(
        [string] $Mailbox,

        [Parameter(Mandatory)]
        [datetime] $StartDate,

        [Parameter(Mandatory)]
        [datetime] $EndDate,

        [int] $LongMeetingDays = 365,

        [int] $MaxModifiedOccurrences = 20,

        [string] $AccessToken,

        [AllowNull()]
        [object[]] $Events
    )

    if ($null -eq $Events) {
        Connect-DiagnosticGraph -AccessToken $AccessToken
        $Events = Get-CalendarEvent -Mailbox $Mailbox -AccessToken $AccessToken
    }

    $inRange = @($Events | Where-Object { Test-EventInRange -CalendarEvent $_ -StartDate $StartDate -EndDate $EndDate })

    Write-Verbose "$($inRange.Count) of $(@($Events).Count) event(s) fall inside the requested range."

    $noEndDate = [System.Collections.Generic.List[object]]::new()
    $longMeetings = [System.Collections.Generic.List[object]]::new()
    $modified = [System.Collections.Generic.List[object]]::new()
    $duplicates = [System.Collections.Generic.List[object]]::new()
    $all = [System.Collections.Generic.List[object]]::new()

    foreach ($calendarEvent in $inRange) {
        $modifiedCount = 0

        if ($null -ne (Get-EventRecurrenceRange -CalendarEvent $calendarEvent)) {
            $modifiedCount = Get-ModifiedOccurrenceCount -CalendarEvent $calendarEvent -StartDate $StartDate -EndDate $EndDate `
                -Mailbox $Mailbox -AccessToken $AccessToken
        }

        $record = ConvertTo-DiagnosticRecord -CalendarEvent $calendarEvent -ModifiedOccurrenceCount $modifiedCount
        $all.Add($record)

        if (Test-EventNoEndDate -CalendarEvent $calendarEvent) { $noEndDate.Add($record) }
        if (Test-EventTooLong -CalendarEvent $calendarEvent -LongMeetingDays $LongMeetingDays) { $longMeetings.Add($record) }
        if ($modifiedCount -gt $MaxModifiedOccurrences) { $modified.Add($record) }
    }

    foreach ($group in (Get-DuplicateICalUid -Events $inRange)) {
        foreach ($duplicate in $group.Group) {
            $duplicates.Add((ConvertTo-DiagnosticRecord -CalendarEvent $duplicate))
        }
    }

    return @{
        Mailbox                = $Mailbox
        StartDate              = $StartDate
        EndDate                = $EndDate
        LongMeetingDays        = $LongMeetingDays
        MaxModifiedOccurrences = $MaxModifiedOccurrences
        TotalEvents            = $inRange.Count
        NoEndDate              = $noEndDate.ToArray()
        LongMeetings           = $longMeetings.ToArray()
        ModifiedOccurrences    = $modified.ToArray()
        DuplicateICalUIDs      = $duplicates.ToArray()
        AllEvents              = $all.ToArray()
    }
}

function Invoke-CalendarDiagnosticReport {
    [CmdletBinding()]
    param(
        [string] $Mailbox,
        [datetime] $StartDate,
        [datetime] $EndDate,
        [string] $OutputPath,
        [int] $LongMeetingDays,
        [int] $MaxModifiedOccurrences,
        [string] $AccessToken
    )

    if ($StartDate -ge $EndDate) {
        throw "StartDate ($StartDate) must be earlier than EndDate ($EndDate)."
    }

    Write-Verbose "Checking the calendar of '$Mailbox' between $StartDate and $EndDate."

    $result = Invoke-CalendarDiagnostic -Mailbox $Mailbox -StartDate $StartDate -EndDate $EndDate `
        -LongMeetingDays $LongMeetingDays -MaxModifiedOccurrences $MaxModifiedOccurrences -AccessToken $AccessToken

    $files = Write-DiagnosticReport -Result $result -OutputPath $OutputPath

    foreach ($line in (Format-DiagnosticCount -Result $result)) {
        Write-Information $line -InformationAction Continue
    }

    Write-Information '' -InformationAction Continue
    Write-Information 'Report files:' -InformationAction Continue
    foreach ($file in $files) { Write-Information "  $file" -InformationAction Continue }

    return $files
}

# Only run when the script is executed, not when it is dot sourced (for tests).
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-CalendarDiagnosticReport -Mailbox $Mailbox -StartDate $StartDate -EndDate $EndDate -OutputPath $OutputPath `
        -LongMeetingDays $LongMeetingDays -MaxModifiedOccurrences $MaxModifiedOccurrences -AccessToken $AccessToken
}
