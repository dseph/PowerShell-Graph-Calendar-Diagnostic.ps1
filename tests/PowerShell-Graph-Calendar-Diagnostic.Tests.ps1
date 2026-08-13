#Requires -Modules Pester

BeforeAll {
    . (Join-Path $PSScriptRoot '..' 'PowerShell-Graph-Calendar-Diagnostic.ps1')

    function New-TestEvent {
        param(
            [string] $Subject = 'Test',
            [string] $ICalUId = 'uid-1',
            [datetime] $Start = ([datetime]'2024-01-01T09:00:00Z'),
            [datetime] $End = ([datetime]'2024-01-01T10:00:00Z'),
            [string] $RangeType,
            [string] $RangeStartDate = '2024-01-01',
            [string] $RangeEndDate = '2024-06-01'
        )

        $testEvent = [pscustomobject] @{
            id         = [guid]::NewGuid().ToString()
            iCalUId    = $ICalUId
            subject    = $Subject
            type       = 'singleInstance'
            start      = [pscustomobject] @{ dateTime = $Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            end        = [pscustomobject] @{ dateTime = $End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = 'UTC' }
            recurrence = $null
            webLink    = 'https://example.invalid'
            organizer  = [pscustomobject] @{ emailAddress = [pscustomobject] @{ address = 'organizer@contoso.com' } }
        }

        if ($RangeType) {
            $testEvent.type = 'seriesMaster'
            $testEvent.recurrence = [pscustomobject] @{
                range = [pscustomobject] @{
                    type      = $RangeType
                    startDate = $RangeStartDate
                    endDate   = $RangeEndDate
                }
            }
        }

        return $testEvent
    }
}

Describe 'Test-EventNoEndDate' {
    It 'reports a recurring series without an end date' {
        Test-EventNoEndDate -CalendarEvent (New-TestEvent -RangeType 'noEnd') | Should -BeTrue
    }

    It 'does not report a series with an end date' {
        Test-EventNoEndDate -CalendarEvent (New-TestEvent -RangeType 'endDate') | Should -BeFalse
    }

    It 'does not report a single instance meeting' {
        Test-EventNoEndDate -CalendarEvent (New-TestEvent) | Should -BeFalse
    }
}

Describe 'Test-EventTooLong' {
    It 'reports a meeting longer than a year' {
        $longEvent = New-TestEvent -Start ([datetime]'2024-01-01T09:00:00Z') -End ([datetime]'2025-06-01T09:00:00Z')
        Test-EventTooLong -CalendarEvent $longEvent -LongMeetingDays 365 | Should -BeTrue
    }

    It 'does not report a one hour meeting' {
        Test-EventTooLong -CalendarEvent (New-TestEvent) -LongMeetingDays 365 | Should -BeFalse
    }
}

Describe 'Get-DuplicateICalUid' {
    It 'groups events that share an iCalUId' {
        $events = @(
            (New-TestEvent -ICalUId 'dup'),
            (New-TestEvent -ICalUId 'dup'),
            (New-TestEvent -ICalUId 'unique')
        )

        $groups = @(Get-DuplicateICalUid -Events $events)
        $groups.Count | Should -Be 1
        $groups[0].Name | Should -Be 'dup'
        $groups[0].Count | Should -Be 2
    }

    It 'returns nothing when every iCalUId is unique' {
        @(Get-DuplicateICalUid -Events @((New-TestEvent -ICalUId 'a'), (New-TestEvent -ICalUId 'b'))).Count | Should -Be 0
    }
}

Describe 'Test-EventInRange' {
    It 'keeps an event inside the range' {
        Test-EventInRange -CalendarEvent (New-TestEvent) -StartDate ([datetime]'2023-01-01Z') -EndDate ([datetime]'2025-01-01Z') |
            Should -BeTrue
    }

    It 'drops an event outside the range' {
        Test-EventInRange -CalendarEvent (New-TestEvent) -StartDate ([datetime]'2020-01-01Z') -EndDate ([datetime]'2021-01-01Z') |
            Should -BeFalse
    }

    It 'keeps a never ending series that started before the range' {
        $seriesEvent = New-TestEvent -RangeType 'noEnd' -RangeStartDate '2010-01-01'
        Test-EventInRange -CalendarEvent $seriesEvent -StartDate ([datetime]'2024-01-01Z') -EndDate ([datetime]'2025-01-01Z') |
            Should -BeTrue
    }
}

Describe 'Invoke-CalendarDiagnostic' {
    BeforeAll {
        Mock Get-ModifiedOccurrenceCount { 25 }

        $script:events = @(
            (New-TestEvent -Subject 'Never ending' -ICalUId 'a' -RangeType 'noEnd'),
            (New-TestEvent -Subject 'Very long' -ICalUId 'b' -Start ([datetime]'2024-01-01T09:00:00Z') -End ([datetime]'2025-06-01T09:00:00Z')),
            (New-TestEvent -Subject 'Duplicate 1' -ICalUId 'dup'),
            (New-TestEvent -Subject 'Duplicate 2' -ICalUId 'dup'),
            (New-TestEvent -Subject 'Healthy' -ICalUId 'c')
        )

        $script:result = Invoke-CalendarDiagnostic -Mailbox 'user@contoso.com' -Events $script:events `
            -StartDate ([datetime]'2023-01-01Z') -EndDate ([datetime]'2026-01-01Z')
    }

    It 'finds each category of issue' {
        @($script:result.NoEndDate).Count | Should -Be 1
        @($script:result.LongMeetings).Count | Should -Be 1
        @($script:result.ModifiedOccurrences).Count | Should -Be 1
        @($script:result.DuplicateICalUIDs).Count | Should -Be 2
        $script:result.TotalEvents | Should -Be 5
    }

    It 'writes a summary and one .csv per check' {
        $outputPath = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())

        try {
            $files = @(Write-DiagnosticReport -Result $script:result -OutputPath $outputPath)

            $files.Count | Should -Be 6
            foreach ($file in $files) { Test-Path -LiteralPath $file | Should -BeTrue }
            @(Get-ChildItem -Path $outputPath -Filter '*.csv').Count | Should -Be 5
            (Get-Content -LiteralPath ($files | Where-Object { $_ -like '*Summary.txt' }) -Raw) |
                Should -Match 'Meetings with no ending date'
        }
        finally {
            Remove-Item -LiteralPath $outputPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
