#Requires -Version 7.0
<#
.SYNOPSIS
    Holiday party awards from your git history (and GitHub PRs).

.DESCRIPTION
    Scans one or more git repositories and produces a leaderboard for ~26 funny
    and serious developer awards, plus GitHub-based awards from merged pull
    requests: Mentor, Quick Draw, The Gatekeeper (most PRs reviewed) and Review
    Dodger (fewest). Every award is calculated separately for each repository,
    so each repo gets its own winners and more people get a shot at first place.

    The AI awards only count things that leave a measurable trace: AI Whisperer
    (commits carrying an AI co-author trailer such as "Co-Authored-By: Claude",
    i.e. the agent wrote the commit), Prompt Engineer (commits that edit AI
    instruction files such as CLAUDE.md, AGENTS.md, .claude/ or .cursor/) and
    Seat Warmer (the GitHub Copilot seat that has gone unused the longest, read
    from the org's seat list, which needs an org owner's gh login). Changes an
    agent made that a person then tested and committed by hand leave no trace
    in git, so no award tries to guess at them.

    The report is a markdown matrix: one row per award, one column per repo
    (plus an "All Repos" column, skip it with -NoCombined), and the top N
    finishers in every cell. Use -Layout Sections for a long per-repo listing
    instead. All times are converted to US Eastern.

    Requires: PowerShell 7+, git on PATH. For Mentor / Quick Draw also the
    GitHub CLI (https://cli.github.com) with `gh auth login` done once.
    The GitHub repo is detected from each local repo's "origin" remote, or pass
    -Github owner/repo. Skip GitHub awards with -NoGithub.

    Tip: author names are merged using .mailmap if the repo has one. Create one
    if the same person appears under several names/emails.

.PARAMETER Repos
    Paths to local git repositories. Each repo gets its own full set of awards.

.PARAMETER Since
    Start of the period. Anything git understands. Default: "1 year ago".

.PARAMETER MinCommits
    Minimum commits (within that repo) to qualify for average-based awards
    (Novelist, Haiku Master, Small Steps).

.PARAMETER Exclude
    Regex matched against git author names, emails and GitHub logins; matches are
    skipped as bots. Default: [bot] accounts, dependabot, github-actions,
    renovate, claude, copilot, compile, amp-autopilot, chatgpt and codex.

.PARAMETER Top
    How many finishers to show per award per repo. Default 3.

.PARAMETER Layout
    Matrix (default): awards as rows, repos as columns, top N in each cell.
    Sections: one heading per repo with every award listed underneath.

.PARAMETER NoCombined
    Only report per-repo awards; skip the "All Repos" column/section.
    Cross-Product Champion and Seat Warmer are only awarded for all repos combined.

.PARAMETER NoCopilot
    Skip the Seat Warmer award (it reads the org's Copilot seat list).

.PARAMETER GhMap
    Map a GitHub login (or GitHub display name) to the name used in git, e.g.
    -GhMap 'jsmith=John Smith'.

.EXAMPLE
    ./Git-Awards.ps1 D:\git\AIMS D:\git\AIMSEvents D:\git\AIMSLicenseGenerator D:\git\AIMSMobile5 D:\git\AIMSMobilePay D:\git\AIMSMobilePay2 D:\git\AIMSMobilePayApp D:\git\AIMSTextMessaging D:\git\AIMSWeb90 D:\git\ConversionExporter -GhMap 'EDC-Collin=Collin Puchta' -Out awards.md

.EXAMPLE
    ./Git-Awards.ps1 .\repo1 -Since '2025-12-01' -AllBranches -MinCommits 10 -GhMap 'jsmith=John Smith'

.EXAMPLE
    ./Git-Awards.ps1 D:\git\AIMS D:\git\AIMSWeb90 -Layout Sections -NoCombined -Top 1 -Out awards.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0, ValueFromRemainingArguments)]
    [string[]]$Repos,
    [string]$Since = '1 year ago',
    [string]$Until,
    [switch]$AllBranches,
    [int]$MinCommits = 5,
    [string]$Exclude = '\[bot\]|dependabot|github-actions|renovate|\bclaude\b|\bcopilot\b|\bcompile\b|amp-autopilot|chatgpt|\bcodex\b',
    [int]$Top = 3,
    [string]$Out,
    [string[]]$Github = @(),
    [switch]$NoGithub,
    [switch]$NoCopilot,
    [string[]]$GhMap = @(),
    [int]$MinPrs = 5,
    [int]$PrLimit = 1000,
    [ValidateSet('Matrix', 'Sections')]
    [string]$Layout = 'Matrix',
    [switch]$NoCombined
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# ----------------------------------------------------------------- setup
$Eastern = $null
foreach ($id in 'America/New_York', 'Eastern Standard Time') {
    try { $Eastern = [TimeZoneInfo]::FindSystemTimeZoneById($id); break } catch { }
}
if (-not $Eastern) {
    Write-Warning 'Could not load the Eastern time zone; using the timezone stored in each commit.'
}

$Fmt = '%x1e%H%x1f%aN%x1f%aE%x1f%aI%x1f%P%x1f%B%x1f'
$RS = [char]0x1e
$US = [char]0x1f
$ic = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

$reNoise   = [regex]::new('(package-lock\.json|yarn\.lock|pnpm-lock\.yaml|composer\.lock|\.min\.(js|css)|(^|/)(node_modules|vendor|dist|build|bin|obj)/|\.(svg|map|dll|exe|png|jpg|gif|pdf)$)', $ic)
$reDoc     = [regex]::new('(\.md$|\.rst$|(^|/)docs?/|readme)', $ic)
$reTest    = [regex]::new('(test|spec|__tests__)', $ic)
$reApology = [regex]::new('\b(oops|whoops|typo|forgot|forgotten|my bad|fix fix|missed|silly|derp|d''oh)\b', $ic)
$reSwear   = [regex]::new('\b(damn|dammit|hell|crap|wtf|shit\w*|fuck\w*|bastard|piss\w*|ass)\b', $ic)
$reBugfix  = [regex]::new('\b(fix(es|ed)?|bug|hotfix|resolve[sd]?|closes?)\b', $ic)
$reTicket  = [regex]::new('\b[A-Z][A-Z0-9]+-\d+\b|#\d+')
$reGeneric = [regex]::new('^(fix|fixes|wip|update|updates|stuff|asdf+|test|tmp|misc|changes?|commit|minor|more|\.+)\W*$', $ic)
$reEmoji   = [regex]::new('[\uD83C-\uD83E][\uDC00-\uDFFF]|[\u2600-\u27BF]|(?<!\w):[a-z][a-z0-9_+-]*:', $ic)
# AI-assisted commit: a co-author trailer naming an AI tool (Claude Code, Copilot,
# Amp, Cursor, aider, ...) or a "Generated with" footer. Matched per line.
$reAI      = [regex]::new('^co-authored-by:.*\b(claude|anthropic|copilot|cursor|aider|chatgpt|openai|gemini|codex|devin|windsurf|amp-autopilot|ampcode)\b|@amp\.local|generated with \[?(claude|copilot|cursor|amp)\b|\bai-generated\b|\(aider\)',
                          $ic -bor [System.Text.RegularExpressions.RegexOptions]::Multiline)
# AI instruction / config files (Claude Code, Codex, Gemini, Cursor, Copilot,
# Windsurf, Cline, Roo, aider, MCP). Third-party copies (vendor/, node_modules/) don't count.
$reAiFiles    = [regex]::new('(^|/)(CLAUDE|AGENTS|GEMINI)\.md$|(^|/)copilot-instructions\.md$|(^|/)\.(cursorrules|windsurfrules|clinerules|roomodes|mcp\.json|aider[^/]*)$|(^|/)\.(claude|cursor|codex|amp|continue|roo|windsurf|gemini|github/(instructions|prompts|agents))/', $ic)
$reThirdParty = [regex]::new('(^|/)([^/]*vendor[^/]*|node_modules|bower_components|third[-_]?party)/', $ic)
$reExclude = if ($Exclude) { [regex]::new($Exclude, $ic) } else { $null }

# ----------------------------------------------------------------- helpers
function Format-Num($x) { ([double]$x).ToString('N0', [cultureinfo]::InvariantCulture) }

function Get-NthWeekday([int]$Year, [int]$Month, [DayOfWeek]$Dow, [int]$N) {
    $d = [datetime]::new($Year, $Month, 1)
    $d = $d.AddDays((([int]$Dow - [int]$d.DayOfWeek) + 7) % 7)
    $d.AddDays(7 * ($N - 1))
}

function Get-LastWeekday([int]$Year, [int]$Month, [DayOfWeek]$Dow) {
    $d = [datetime]::new($Year, $Month, 1).AddMonths(1).AddDays(-1)
    $d.AddDays(-((([int]$d.DayOfWeek - [int]$Dow) + 7) % 7))
}

$script:HolidayCache = @{}
function Get-Holidays([int]$Year) {
    if (-not $script:HolidayCache.ContainsKey($Year)) {
        $k = 'yyyy-MM-dd'
        $script:HolidayCache[$Year] = @{
            ([datetime]::new($Year, 1, 1).ToString($k))                       = "New Year's Day"
            ((Get-LastWeekday $Year 5 Monday).ToString($k))                   = 'Memorial Day'
            ([datetime]::new($Year, 7, 4).ToString($k))                       = 'July 4th'
            ((Get-NthWeekday $Year 9 Monday 1).ToString($k))                  = 'Labor Day'
            ((Get-NthWeekday $Year 11 Thursday 4).ToString($k))               = 'Thanksgiving'
            ([datetime]::new($Year, 12, 24).ToString($k))                     = 'Christmas Eve'
            ([datetime]::new($Year, 12, 25).ToString($k))                     = 'Christmas Day'
            ([datetime]::new($Year, 12, 31).ToString($k))                     = "New Year's Eve"
        }
    }
    $script:HolidayCache[$Year]
}

function Test-Weekend([datetime]$d) { $d.DayOfWeek -in 'Saturday', 'Sunday' }

function Get-NextBusinessDay([datetime]$d) {
    $d = $d.AddDays(1)
    while (Test-Weekend $d) { $d = $d.AddDays(1) }
    $d
}

function Get-LongestStreak($DateStrings) {
    $days = @($DateStrings | ForEach-Object { [datetime]::ParseExact($_, 'yyyy-MM-dd', $null) } |
              Where-Object { -not (Test-Weekend $_) } | Sort-Object)
    if ($days.Count -eq 0) { return 0 }
    $best = $cur = 1
    for ($i = 1; $i -lt $days.Count; $i++) {
        if ((Get-NextBusinessDay $days[$i - 1]) -eq $days[$i]) { $cur++ } else { $cur = 1 }
        if ($cur -gt $best) { $best = $cur }
    }
    $best
}

function Get-Median($Values) {
    $s = @($Values | Sort-Object)
    $n = $s.Count
    if ($n -eq 0) { return 0 }
    if ($n % 2) { $s[[int](($n - 1) / 2)] } else { ($s[$n / 2 - 1] + $s[$n / 2]) / 2 }
}

function Format-Duration([double]$Hours) {
    if ($Hours -lt 48) { '{0:N1} hrs' -f $Hours } else { '{0:N1} days' -f ($Hours / 24) }
}

function New-Stat {
    $h = @{}
    foreach ($k in 'commits', 'merges', 'conflict_merges', 'after5', 'vampire', 'early', 'weekend',
                   'friday_late', 'holiday', 'msg_len', 'subj_len', 'generic', 'apology', 'reverts',
                   'bugfix', 'emoji', 'swear', 'ai', 'ai_files', 'added', 'deleted', 'files', 'docs', 'tests', 'big') {
        $h[$k] = 0
    }
    $h
}

# A short display name per repo path: the folder name, or the full path when
# two repos share a folder name.
function Get-RepoLabels([string[]]$Paths) {
    $leaf = @{}
    foreach ($p in $Paths) {
        $full = $p
        try { $full = (Resolve-Path -LiteralPath $p -ErrorAction Stop).Path } catch { }
        $l = Split-Path -Leaf ($full.TrimEnd('\', '/'))
        $leaf[$p] = if ($l) { $l } else { $full }
    }
    $counts = @{}
    foreach ($l in $leaf.Values) { $counts[$l] = 1 + [int]$counts[$l] }
    $labels = @{}
    foreach ($p in $Paths) { $labels[$p] = if ($counts[$leaf[$p]] -gt 1) { $p } else { $leaf[$p] } }
    $labels
}

# ----------------------------------------------------------------- git
function Get-RepoCommits([string]$Repo) {
    $a = @('-C', $Repo, 'log', "--since=$Since", "--format=$Fmt", '--numstat')
    if ($Until) { $a += "--until=$Until" }
    if ($AllBranches) { $a += '--all' }
    $raw = & git @a 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Skipping ${Repo}: $(($raw | ForEach-Object { "$_" }) -join ' ')"
        return
    }
    $text = ($raw | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } |
             ForEach-Object { "$_" }) -join "`n"
    foreach ($rec in ($text.Split($RS) | Select-Object -Skip 1)) {
        $p = $rec.Split($US)
        if ($p.Count -lt 6) { continue }
        [pscustomobject]@{
            Sha     = $p[0]
            Name    = $p[1]
            Email   = $p[2]
            Iso     = $p[3]
            Parents = @($p[4].Trim() -split '\s+' | Where-Object { $_ })
            Msg     = $p[5].Trim()
            Stat    = if ($p.Count -gt 6) { $p[6] } else { '' }
        }
    }
}

# ----------------------------------------------------------------- stat buckets
# A "bucket" holds per-author stats for one scope: a single repo, or all repos
# combined. Each commit is measured once and then added to both buckets.
function New-Bucket {
    @{ Stats = @{}; RepoSets = @{}; DateSets = @{}; HolidayKinds = @{}; Total = 0 }
}

function Measure-Commit($c) {
    $msg = $c.Msg
    $subject = ($msg -split "`r?`n")[0]
    $dto = [DateTimeOffset]::Parse($c.Iso, [cultureinfo]::InvariantCulture)
    $t = if ($Eastern) { [TimeZoneInfo]::ConvertTime($dto, $Eastern) } else { $dto }

    if ($c.Parents.Count -gt 1) {
        return @{ IsMerge = $true; Conflict = $msg.ToLowerInvariant().Contains('conflict') }
    }

    $dateKey = $t.Date.ToString('yyyy-MM-dd')
    $h = $t.Hour
    $dow = $t.DayOfWeek
    $hol = (Get-Holidays $t.Year)[$dateKey]

    $inc = @{
        commits     = 1
        after5      = [int]($h -ge 17)
        vampire     = [int]($h -lt 5)
        early       = [int]($h -ge 5 -and $h -lt 7)
        weekend     = [int]($dow -in 'Saturday', 'Sunday')
        friday_late = [int]($dow -eq 'Friday' -and $h -ge 16)
        holiday     = [int][bool]$hol
        msg_len     = $msg.Length
        subj_len    = $subject.Length
        generic     = [int]($reGeneric.IsMatch($subject.Trim()))
        apology     = [int]($reApology.IsMatch($msg))
        reverts     = [int]($subject.ToLowerInvariant().StartsWith('revert'))
        bugfix      = [int]($reBugfix.IsMatch($msg) -or $reTicket.IsMatch($msg))
        emoji       = [int]($reEmoji.IsMatch($msg))
        swear       = $reSwear.Matches($msg).Count
        ai          = [int]($reAI.IsMatch($msg))
        ai_files    = 0
        added       = 0
        deleted     = 0
        files       = 0
        docs        = 0
        tests       = 0
    }
    foreach ($line in ($c.Stat -split "`n")) {
        $bits = $line.Trim("`r").Split("`t")
        if ($bits.Count -ne 3) { continue }
        $path = $bits[2]
        if ($reDoc.IsMatch($path)) { $inc.docs = 1 }
        if ($reTest.IsMatch($path)) { $inc.tests = 1 }
        if ($reAiFiles.IsMatch($path) -and -not $reThirdParty.IsMatch($path)) { $inc.ai_files = 1 }
        if ($reNoise.IsMatch($path) -or $bits[0] -eq '-') { continue }
        $inc.added += [int]$bits[0]
        $inc.deleted += [int]$bits[1]
        $inc.files++
    }

    @{ IsMerge = $false; Inc = $inc; Big = $inc.added + $inc.deleted; DateKey = $dateKey; Holiday = $hol }
}

function Add-Commit($b, [string]$Name, [string]$RepoLabel, $m) {
    if (-not $b.Stats.ContainsKey($Name)) {
        $b.Stats[$Name] = New-Stat
        $b.RepoSets[$Name] = [System.Collections.Generic.HashSet[string]]::new()
        $b.DateSets[$Name] = [System.Collections.Generic.HashSet[string]]::new()
        $b.HolidayKinds[$Name] = [System.Collections.Generic.HashSet[string]]::new()
    }
    $s = $b.Stats[$Name]
    if ($m.IsMerge) {
        $s.merges++
        if ($m.Conflict) { $s.conflict_merges++ }
        return
    }
    $b.Total++
    foreach ($k in $m.Inc.Keys) { $s[$k] += $m.Inc[$k] }
    if ($m.Big -gt $s.big) { $s.big = $m.Big }
    [void]$b.RepoSets[$Name].Add($RepoLabel)
    [void]$b.DateSets[$Name].Add($m.DateKey)
    if ($m.Holiday) { [void]$b.HolidayKinds[$Name].Add($m.Holiday) }
}

function Complete-Bucket($b) {
    foreach ($name in @($b.Stats.Keys)) {
        $s = $b.Stats[$name]
        $s.repos = $b.RepoSets[$name].Count
        $s.streak = Get-LongestStreak $b.DateSets[$name]
        $s.holiday_kinds = $b.HolidayKinds[$name].Count
        $s.lines = $s.added + $s.deleted
        $n = [double]$s.commits
        $s.avg_msg  = if ($n) { $s.msg_len / $n } else { 0 }
        $s.avg_subj = if ($n) { $s.subj_len / $n } else { 0 }
        $s.avg_size = if ($n) { $s.lines / $n } else { 0 }
        $s.ai_share = if ($n) { $s.ai / $n } else { 0 }
    }
}

# ----------------------------------------------------------------- github
function Get-GitSinceTimestamp([string]$Repo) {
    try {
        $o = & git -C $Repo rev-parse "--since=$Since" 2>$null
        if ($LASTEXITCODE -eq 0 -and "$o" -match '--max-age=(\d+)') { return [long]$Matches[1] }
    } catch { }
    $null
}

function Get-GitHubSlug([string]$Repo) {
    try {
        $url = (& git -C $Repo remote get-url origin 2>$null | Out-String).Trim()
        if ($url -match 'github\.com[:/]+([^/]+)/([^/]+?)(?:\.git)?/?$') { return "$($Matches[1])/$($Matches[2])" }
    } catch { }
    $null
}

function Invoke-Gh([string[]]$Argv) {
    $raw = & gh @Argv 2>&1
    if ($LASTEXITCODE -ne 0) { throw (($raw | ForEach-Object { "$_" }) -join "`n") }
    ($raw | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } |
     ForEach-Object { "$_" }) -join "`n"
}

# ConvertFrom-Json turns ISO date strings into DateTime; handle both forms.
function ConvertTo-UtcDate($v) {
    if ($v -is [datetime]) { return $v.ToUniversalTime() }
    [DateTimeOffset]::Parse("$v", [cultureinfo]::InvariantCulture).UtcDateTime
}

function Test-GhSkip($Login) { (-not $Login) -or ($reExclude -and $reExclude.IsMatch($Login)) }

# Per-person GitHub record: review submissions, inline comments, open-to-merge
# hours of their own merged PRs, and the distinct PRs ("owner/repo#n") they reviewed.
function New-GhRecord {
    @{ reviews = 0; inline = 0
       hours = [System.Collections.Generic.List[double]]::new()
       reviewed = [System.Collections.Generic.HashSet[string]]::new()
       has_seat = $false; seat_last = $null; seat_editor = '' }
}

function Get-GhUser($Users, [string]$Login) {
    if (-not $Users.ContainsKey($Login)) { $Users[$Login] = New-GhRecord }
    $Users[$Login]
}

# Returns @{ BySlug = [ordered]@{ 'owner/repo' = @{ login = @{reviews; inline; hours} } }; Names = @{ login = display name } }
function Get-GitHubData([string[]]$Slugs, [long]$SinceTs) {
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        Write-Warning 'GitHub CLI (gh) not found; skipping GitHub awards. See https://cli.github.com'
        return $null
    }
    $sinceDt = [DateTimeOffset]::FromUnixTimeSeconds($SinceTs).UtcDateTime
    $bySlug = [ordered]@{}
    $names = @{}

    foreach ($slug in $Slugs) {
        Write-Host "Fetching GitHub PRs for $slug ..."
        try {
            $prs = Invoke-Gh @('pr', 'list', '--repo', $slug, '--state', 'merged',
                '--search', "merged:>=$($sinceDt.ToString('yyyy-MM-dd'))",
                '--limit', "$PrLimit", '--json', 'number,author,createdAt,mergedAt,reviews') | ConvertFrom-Json
            $inlineRaw = Invoke-Gh @('api', '--paginate',
                "repos/$slug/pulls/comments?since=$($sinceDt.ToString('yyyy-MM-ddTHH:mm:ssZ'))&per_page=100",
                '--jq', '.[] | {u: .user.login, p: .pull_request_url, c: .created_at}')
        } catch {
            Write-Warning "gh failed for $slug (run 'gh auth login'?): $($_.Exception.Message)"
            continue
        }

        $users = @{}
        $prAuthor = @{}
        foreach ($pr in @($prs)) {
            $login = $pr.author.login
            $prAuthor[[int]$pr.number] = $login
            if ($login -and $pr.author.name) { $names[$login] = $pr.author.name }
            if ($pr.author.is_bot -or (Test-GhSkip $login)) { continue }
            if ($pr.mergedAt -and $pr.createdAt) {
                $hrs = ((ConvertTo-UtcDate $pr.mergedAt) - (ConvertTo-UtcDate $pr.createdAt)).TotalHours
                (Get-GhUser $users $login).hours.Add($hrs)
            }
            foreach ($r in @($pr.reviews)) {
                $rl = $r.author.login
                if ((Test-GhSkip $rl) -or $rl -eq $login -or $r.state -eq 'PENDING') { continue }
                $u = Get-GhUser $users $rl
                $u.reviews++
                [void]$u.reviewed.Add("$slug#$($pr.number)")
            }
        }

        foreach ($line in ($inlineRaw -split "`n")) {
            if (-not $line.Trim()) { continue }
            try {
                $c = $line | ConvertFrom-Json
                $num = [int]($c.p.TrimEnd('/').Split('/')[-1])
                $created = ConvertTo-UtcDate $c.c
            } catch { continue }
            if ($created -lt $sinceDt -or -not $prAuthor.ContainsKey($num)) { continue }
            if ((Test-GhSkip $c.u) -or $c.u -eq $prAuthor[$num]) { continue }
            $u = Get-GhUser $users $c.u
            $u.inline++
            [void]$u.reviewed.Add("$slug#$num")
        }
        $bySlug[$slug] = $users
    }

    # Copilot seats: login -> @{ last = [datetime] (UTC) or $null; editor }. Needs an org owner.
    $seats = @{}
    if (-not $NoCopilot) {
        foreach ($owner in @($Slugs | ForEach-Object { ($_ -split '/')[0] } | Select-Object -Unique)) {
            Write-Host "Fetching Copilot seats for $owner ..."
            try {
                $raw = Invoke-Gh @('api', '--paginate', "orgs/$owner/copilot/billing/seats?per_page=100",
                    '--jq', '.seats[] | {l: .assignee.login, a: .last_activity_at, e: .last_activity_editor}')
            } catch {
                Write-Warning "Could not read Copilot seats for $owner (needs an org owner's gh login); skipping Seat Warmer: $($_.Exception.Message)"
                continue
            }
            foreach ($line in ($raw -split "`n")) {
                if (-not $line.Trim()) { continue }
                try { $seat = $line | ConvertFrom-Json } catch { continue }
                if (-not $seat.l -or (Test-GhSkip $seat.l)) { continue }
                $seats[$seat.l] = @{ last = $(if ($seat.a) { ConvertTo-UtcDate $seat.a } else { $null }); editor = "$($seat.e)" }
            }
        }
    }
    @{ BySlug = $bySlug; Names = $names; Seats = $seats }
}

# GitHub login -> report name: -GhMap first, then the PR author's display name,
# then a case-insensitive match against the git author names.
function Resolve-GhName([string]$Login, $Names, $Override, $Lower) {
    if ($Override.ContainsKey($Login)) { return $Override[$Login] }
    $hint = if ($Names.ContainsKey($Login)) { $Names[$Login] } else { '' }
    if ($hint -and $Override.ContainsKey($hint)) { return $Override[$hint] }
    $key = $(if ($hint) { $hint } else { $Login }).ToLowerInvariant()
    if ($Lower.ContainsKey($key)) { $Lower[$key] } elseif ($hint) { $hint } else { $Login }
}

# Attach Copilot seat info to a merged user table. Only people with PR activity
# or commits in the scanned repos are added, so the report never lists the whole
# seat roster.
function Add-GhSeats($Merged, $Seats, $Names, $Override, $Lower, $Committers) {
    foreach ($login in $Seats.Keys) {
        $display = Resolve-GhName $login $Names $Override $Lower
        if (-not ($Merged.ContainsKey($display) -or $Committers.ContainsKey($display))) { continue }
        if (-not $Merged.ContainsKey($display)) { $Merged[$display] = New-GhRecord }
        $m = $Merged[$display]
        $m.has_seat = $true
        $m.seat_last = $Seats[$login].last
        $m.seat_editor = $Seats[$login].editor
    }
}

# Combine one or more per-slug user tables, mapping GitHub logins to git author names.
function Merge-GhUsers($UserTables, $Names, $Override, $Lower) {
    $merged = @{}
    foreach ($users in @($UserTables)) {
        foreach ($login in $users.Keys) {
            $u = $users[$login]
            $display = Resolve-GhName $login $Names $Override $Lower
            if (-not $merged.ContainsKey($display)) { $merged[$display] = New-GhRecord }
            $mm = $merged[$display]
            $mm.reviews += $u.reviews
            $mm.inline += $u.inline
            $mm.hours.AddRange($u.hours)
            $mm.reviewed.UnionWith($u.reviewed)
        }
    }
    $merged
}

# ----------------------------------------------------------------- collect
$labels = Get-RepoLabels $Repos
$all = New-Bucket                 # every repo combined (commits de-duplicated by SHA)
$repoBuckets = [ordered]@{}       # label -> bucket for that repo alone
$seen = [System.Collections.Generic.HashSet[string]]::new()

foreach ($repo in $Repos) {
    $label = $labels[$repo]
    $b = New-Bucket
    $repoBuckets[$label] = $b
    foreach ($c in (Get-RepoCommits $repo)) {
        if ($reExclude -and ($reExclude.IsMatch($c.Name) -or $reExclude.IsMatch($c.Email))) { continue }
        $mc = Measure-Commit $c
        Add-Commit $b $c.Name $label $mc
        if ($seen.Add($c.Sha)) { Add-Commit $all $c.Name $label $mc }
    }
    Complete-Bucket $b
}
Complete-Bucket $all

if ($all.Stats.Count -eq 0) { Write-Error 'No commits found.'; return }

# ----------------------------------------------------------------- awards
# Cat, Title, Blurb, Score, Disp, Min (int or 'min'), High, CombinedOnly
function New-Award($Cat, $Title, $Blurb, $Score, $Disp, $Min, $High, [bool]$CombinedOnly = $false) {
    [pscustomobject]@{ Cat = $Cat; Title = $Title; Blurb = $Blurb; Score = $Score; Disp = $Disp; Min = $Min; High = $High; CombinedOnly = $CombinedOnly }
}

$awards = @(
    New-Award 'Serious' 'Commit King/Queen' 'Most commits' { param($s) $s.commits } { param($s) "$(Format-Num $s.commits) commits" } 1 $true
    New-Award 'Serious' 'Biggest Impact' 'Most lines changed (lockfiles/minified/binaries ignored)' { param($s) $s.lines } { param($s) "$(Format-Num $s.lines) lines" } 1 $true
    New-Award 'Serious' 'Bug Slayer' 'Most commits mentioning fixes or ticket numbers' { param($s) $s.bugfix } { param($s) "$(Format-Num $s.bugfix) fix commits" } 1 $true
    New-Award 'Serious' 'Cross-Product Champion' 'Contributed to the most repositories' { param($s) $s.repos * 1000000 + $s.commits } { param($s) "$(Format-Num $s.repos) repos" } 1 $true $true
    New-Award 'Serious' 'Best Documenter' 'Most commits touching docs/README/markdown' { param($s) $s.docs } { param($s) "$(Format-Num $s.docs) doc commits" } 1 $true
    New-Award 'Serious' 'Test Champion' 'Most commits touching test files' { param($s) $s.tests } { param($s) "$(Format-Num $s.tests) test commits" } 1 $true
    New-Award 'Serious' 'Mr./Ms. Consistency' 'Longest streak of consecutive business days with a commit' { param($s) $s.streak } { param($s) "$(Format-Num $s.streak) business days" } 1 $true
    New-Award 'Serious' 'Small Steps' 'Smallest average commit (small PRs are easy to review)' { param($s) $s.avg_size } { param($s) '{0:N0} lines/commit avg' -f $s.avg_size } 'min' $false
    New-Award 'Funny' 'Works After 5pm' 'Most commits after 5pm Eastern' { param($s) $s.after5 } { param($s) "$(Format-Num $s.after5) commits after 5pm" } 1 $true
    New-Award 'Funny' 'Vampire' 'Most commits between midnight and 5am' { param($s) $s.vampire } { param($s) "$(Format-Num $s.vampire) commits in the dead of night" } 1 $true
    New-Award 'Funny' 'Early Bird' 'Most commits between 5am and 7am' { param($s) $s.early } { param($s) "$(Format-Num $s.early) pre-7am commits" } 1 $true
    New-Award 'Funny' 'Weekend Warrior' 'Most commits on Saturday/Sunday' { param($s) $s.weekend } { param($s) "$(Format-Num $s.weekend) weekend commits" } 1 $true
    New-Award 'Funny' 'Holiday Hero' 'Most commits on a US holiday' { param($s) $s.holiday * 1000 + $s.holiday_kinds } { param($s) "$(Format-Num $s.holiday) holiday commits" } 1 $true
    New-Award 'Funny' 'Friday 4pm Deployer' 'Most commits after 4pm on a Friday' { param($s) $s.friday_late } { param($s) "$(Format-Num $s.friday_late) risky Fridays" } 1 $true
    New-Award 'Funny' 'The Eraser' 'Most lines deleted' { param($s) $s.deleted } { param($s) "$(Format-Num $s.deleted) lines deleted" } 1 $true
    New-Award 'Funny' 'The Novelist' 'Longest average commit message' { param($s) $s.avg_msg } { param($s) '{0:N0} chars avg' -f $s.avg_msg } 'min' $true
    New-Award 'Funny' 'Haiku Master' 'Shortest average commit subject' { param($s) $s.avg_subj } { param($s) '{0:N0} chars avg' -f $s.avg_subj } 'min' $false
    New-Award 'Funny' 'Commit Message Minimalist' "Most 'fix' / 'wip' / 'update' / 'asdf' style messages" { param($s) $s.generic } { param($s) "$(Format-Num $s.generic) cryptic messages" } 1 $true
    New-Award 'Funny' 'The Apologist' 'Most commits with oops/typo/forgot/missed' { param($s) $s.apology } { param($s) "$(Format-Num $s.apology) 'oops' commits" } 1 $true
    New-Award 'Funny' 'Undo Button' 'Most revert commits' { param($s) $s.reverts } { param($s) "$(Format-Num $s.reverts) reverts" } 1 $true
    New-Award 'Funny' 'Merge Conflict Magnet' 'Most merge commits mentioning conflicts' { param($s) $s.conflict_merges } { param($s) "$(Format-Num $s.conflict_merges) conflict merges" } 1 $true
    New-Award 'Funny' 'Emoji Enthusiast' 'Most commits containing emoji or :shortcodes:' { param($s) $s.emoji } { param($s) "$(Format-Num $s.emoji) emoji commits" } 1 $true
    New-Award 'Funny' 'Swear Jar' 'Most colorful language in commit messages' { param($s) $s.swear } { param($s) "$(Format-Num $s.swear) naughty words" } 1 $true
    New-Award 'Funny' 'Big Bang' 'Largest single commit' { param($s) $s.big } { param($s) "$(Format-Num $s.big) lines in one commit" } 1 $true
    New-Award 'Funny' 'AI Whisperer' 'Most commits written by an AI agent (co-author trailers such as "Co-Authored-By: Claude")' { param($s) $s.ai } { param($s) "$(Format-Num $s.ai) agent-written commits, $([math]::Round($s.ai_share * 100))% of theirs" } 1 $true
    New-Award 'Funny' 'Prompt Engineer' 'Most commits editing AI instruction files (CLAUDE.md, AGENTS.md, .claude/, .cursor/, copilot-instructions.md, ...)' { param($s) $s.ai_files } { param($s) "$(Format-Num $s.ai_files) commits teaching the AI" } 1 $true
)

$ghAwards = @(
    [pscustomobject]@{
       Cat   = 'GitHub'
       Title = 'Mentor Award'
       Blurb = "Most PR reviews + inline review comments on teammates' PRs"
       Ok    = { param($u) $true }
       Score = { param($u) $u.reviews + $u.inline }
       Disp  = { param($u) "$(Format-Num $u.reviews) reviews + $(Format-Num $u.inline) inline comments" }
       High  = $true },
    [pscustomobject]@{
       Cat   = 'GitHub'
       Title = 'Quick Draw'
       Blurb = "Fastest median PR turnaround, open to merge (min $MinPrs merged PRs)"
       Ok    = { param($u) $u.hours.Count -ge $MinPrs }
       Score = { param($u) Get-Median $u.hours }
       Disp  = { param($u) "$(Format-Duration (Get-Median $u.hours)) median over $($u.hours.Count) PRs" }
       High  = $false },
    [pscustomobject]@{
       Cat   = 'GitHub'
       Title = 'The Gatekeeper'
       Blurb = "Reviewed the most of other people's pull requests"
       Ok    = { param($u) $true }
       Score = { param($u) $u.reviewed.Count }
       Disp  = { param($u) "$(Format-Num $u.reviewed.Count) PRs reviewed" }
       High  = $true },
    [pscustomobject]@{
       Cat   = 'GitHub'
       Title = 'Review Dodger'
       Blurb = "Reviewed the fewest pull requests while getting their own merged (min $MinPrs merged PRs; ties go to whoever had more merged)"
       Ok    = { param($u) $u.hours.Count -ge $MinPrs }
       Score = { param($u) $u.reviewed.Count * 100000 - $u.hours.Count }
       Disp  = { param($u) "$(Format-Num $u.reviewed.Count) PRs reviewed vs $(Format-Num $u.hours.Count) of their own merged" }
       High  = $false },
    [pscustomobject]@{
       Cat   = 'GitHub'
       Title = 'Seat Warmer'
       Blurb = 'Copilot seat that has gone unused the longest (people with commits or PR activity here only; a never-used seat ranks first)'
       Ok    = { param($u) $u.has_seat }
       Score = { param($u) if ($u.seat_last) { ([datetime]::UtcNow - $u.seat_last).TotalDays } else { 1e6 } }
       Disp  = { param($u)
                 if (-not $u.seat_last) { return 'never used their Copilot seat' }
                 $editor = ($u.seat_editor -split '/')[0]
                 "last Copilot activity $($u.seat_last.ToString('yyyy-MM-dd'))$(if ($editor) { " in $editor" }), $(Format-Num ([datetime]::UtcNow - $u.seat_last).TotalDays) days ago" }
       High  = $true
       CombinedOnly = $true }
)

$medals = '1st', '2nd', '3rd'
function Get-Medal([int]$i) { if ($i -lt 3) { $medals[$i] } else { "$($i + 1)th" } }

# winners: author name -> ordered (award title -> list of columns where they came first)
$winners = @{}
function Add-Winner($Name, $Title, $Context) {
    if (-not $winners.ContainsKey($Name)) { $winners[$Name] = [ordered]@{} }
    $w = $winners[$Name]
    if (-not $w.Contains($Title)) { $w[$Title] = [System.Collections.Generic.List[string]]::new() }
    $w[$Title].Add($Context)
}
function Get-WinCount($w) { $n = 0; foreach ($l in $w.Values) { $n += $l.Count }; $n }

# Top-N for one git award in one bucket. Emits objects with Name and Text (the
# display string), or nothing; callers collect with @(...).
function Get-GitRanking($b, $a) {
    $floor = if ($a.Min -eq 'min') { $MinCommits } else { [int]$a.Min }
    $pool = foreach ($name in $b.Stats.Keys) {
        $s = $b.Stats[$name]
        if ($s.commits -lt $floor) { continue }
        $score = [double](& $a.Score $s)
        if ($a.High -and $score -eq 0) { continue }
        [pscustomobject]@{ Name = $name; Stat = $s; Score = $score }
    }
    $pool = @($pool)
    if ($pool.Count -eq 0) { return }
    $pool | Sort-Object @{ Expression = 'Score'; Descending = $a.High }, Name | Select-Object -First $Top |
        ForEach-Object { [pscustomobject]@{ Name = $_.Name; Text = (& $a.Disp $_.Stat) } }
}

# Top-N for one GitHub award in one merged user table.
function Get-GhRanking($Merged, $ga) {
    $pool = foreach ($name in $Merged.Keys) {
        $u = $Merged[$name]
        if (-not (& $ga.Ok $u)) { continue }
        $score = [double](& $ga.Score $u)
        if ($ga.High -and $score -eq 0) { continue }
        [pscustomobject]@{ Name = $name; U = $u; Score = $score }
    }
    $pool = @($pool)
    if ($pool.Count -eq 0) { return }
    $pool | Sort-Object @{ Expression = 'Score'; Descending = $ga.High }, Name | Select-Object -First $Top |
        ForEach-Object { [pscustomobject]@{ Name = $_.Name; Text = (& $ga.Disp $_.U) } }
}

# ---- GitHub data (fetched up front so each repo column can include its PR awards)
$gh = $null
$override = @{}
$lower = @{}
$repoSlugs = @{}
if (-not $NoGithub) {
    foreach ($repo in $Repos) { $repoSlugs[$labels[$repo]] = Get-GitHubSlug $repo }
    $slugs = if ($Github.Count) { @($Github) }
             else { @($repoSlugs.Values | Where-Object { $_ } | Select-Object -Unique) }
    $ts = $Repos | ForEach-Object { Get-GitSinceTimestamp $_ } | Where-Object { $_ } | Select-Object -First 1

    if ($slugs.Count -eq 0) {
        Write-Warning 'No GitHub remotes detected; skipping Mentor/Quick Draw (use -Github owner/repo).'
    } elseif (-not $ts) {
        Write-Warning 'Could not interpret -Since for GitHub; skipping PR awards.'
    } else {
        $gh = Get-GitHubData $slugs ([long]$ts)
        if ($gh) {
            foreach ($m in $GhMap) {
                $kv = $m -split '=', 2
                if ($kv.Count -eq 2) { $override[$kv[0]] = $kv[1] }
            }
            foreach ($k in $all.Stats.Keys) { $lower[$k.ToLowerInvariant()] = $k }
        }
    }
}

# ---- columns: one per repo with commits, GitHub-only extras, then All Repos
# Each column: Label, Bucket (git stats or $null), Gh (merged GitHub users or $null), Slug, Note, Combined
$columns = [System.Collections.Generic.List[object]]::new()
$usedSlugs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$emptyRepos = [System.Collections.Generic.List[string]]::new()

foreach ($label in $repoBuckets.Keys) {
    $b = $repoBuckets[$label]
    if ($b.Stats.Count -eq 0) { $emptyRepos.Add($label); continue }
    $ghUsers = $null
    $slug = $null
    if ($gh -and $repoSlugs[$label]) {
        $slugKey = @($gh.BySlug.Keys | Where-Object { $_ -eq $repoSlugs[$label] } | Select-Object -First 1)
        if ($slugKey.Count) {
            $slug = $slugKey[0]
            [void]$usedSlugs.Add($slug)
            $ghUsers = Merge-GhUsers $gh.BySlug[$slug] $gh.Names $override $lower
        }
    }
    $columns.Add(@{ Label = $label; Bucket = $b; Gh = $ghUsers; Slug = $slug; Combined = $false
                    Note = "$(Format-Num $b.Total) commits, $($b.Stats.Count) authors" })
}
if ($gh) {
    # repos given with -Github that match no local repo's origin: GitHub awards only
    foreach ($slug in $gh.BySlug.Keys) {
        if ($usedSlugs.Contains($slug)) { continue }
        $columns.Add(@{ Label = $slug; Bucket = $null; Slug = $slug; Combined = $false; Note = 'GitHub only'
                        Gh = (Merge-GhUsers $gh.BySlug[$slug] $gh.Names $override $lower) })
    }
}
if (-not $NoCombined) {
    $allGh = if ($gh) { Merge-GhUsers $gh.BySlug.Values $gh.Names $override $lower } else { $null }
    if ($null -ne $allGh -and $gh.Seats.Count) { Add-GhSeats $allGh $gh.Seats $gh.Names $override $lower $all.Stats }
    $columns.Add(@{ Label = 'All Repos'; Bucket = $all; Gh = $allGh; Slug = $null; Combined = $true
                    Note = "$(Format-Num $all.Total) commits, $($all.Stats.Count) authors" })
}

# ---- rank every award in every column (this also records the #1 finishes)
# Row: Award, Cells (ordered: column label -> ranking array, or $null when the column has no data source), Any
$gitRows = @(foreach ($a in $awards) {
    $cells = [ordered]@{}
    $any = $false
    foreach ($col in $columns) {
        $r = $null
        if ($col.Bucket -and -not ($a.CombinedOnly -and -not $col.Combined)) { $r = @(Get-GitRanking $col.Bucket $a) }
        $cells[$col.Label] = $r
        if ($r -and $r.Count) { $any = $true; Add-Winner $r[0].Name $a.Title $col.Label }
    }
    @{ Award = $a; Cells = $cells; Any = $any }
})
$ghRows = @()
if ($gh) {
    $ghRows = @(foreach ($ga in $ghAwards) {
        $cells = [ordered]@{}
        $any = $false
        foreach ($col in $columns) {
            $r = $null
            if ($null -ne $col.Gh -and -not ($ga.CombinedOnly -and -not $col.Combined)) { $r = @(Get-GhRanking $col.Gh $ga) }
            $cells[$col.Label] = $r
            if ($r -and $r.Count) { $any = $true; Add-Winner $r[0].Name $ga.Title $col.Label }
        }
        @{ Award = $ga; Cells = $cells; Any = $any }
    })
}

# ----------------------------------------------------------------- report
$report = [System.Collections.Generic.List[string]]::new()
$untilLabel = if ($Until) { $Until } else { 'today' }
$report.Add("# Git Awards ($Since to $untilLabel)")
$report.Add("$(Format-Num $all.Total) commits from $($all.Stats.Count) authors across $($Repos.Count) repo(s). Times are US Eastern.")

function ConvertTo-CellText([string]$s) { $s.Replace('|', '\|') }

function Format-Cell($ranked) {
    if ($null -eq $ranked -or $ranked.Count -eq 0) { return '–' }
    $parts = for ($i = 0; $i -lt $ranked.Count; $i++) {
        $name = ConvertTo-CellText $ranked[$i].Name
        $text = ConvertTo-CellText $ranked[$i].Text
        if ($i -eq 0) { "$($i + 1). **$name** ($text)" } else { "$($i + 1). $name ($text)" }
    }
    $parts -join '<br>'
}

# One markdown table: a row per award, a column per entry of $Cols, top N in each cell.
function Write-MatrixTable([string]$Heading, $Rows, $Cols) {
    $rows = @($Rows | Where-Object { $_.Any })
    $cols = @($Cols)
    if ($rows.Count -eq 0 -or $cols.Count -eq 0) { return }
    $report.Add("## $Heading"); $report.Add('')
    $hdr = @('Award') + @($cols | ForEach-Object { "$($_.Label)<br>_$($_.Note)_" })
    $report.Add('| ' + ($hdr -join ' | ') + ' |')
    $report.Add('|' + ((@('---') * $hdr.Count) -join '|') + '|')
    foreach ($row in $rows) {
        $a = $row.Award
        $line = @("**$($a.Title)**<br>_$($a.Blurb)_") + @($cols | ForEach-Object { Format-Cell $row.Cells[$_.Label] })
        $report.Add('| ' + ($line -join ' | ') + ' |')
    }
    $report.Add('')
}

function Write-MatrixLayout {
    $report.Add("Each award is scored separately per repo. Cells show the top $Top, winner in bold.")
    if ($emptyRepos.Count) { $report.Add("No commits in this period: $($emptyRepos -join ', ').") }
    $report.Add('')
    $gitCols = @($columns | Where-Object { $_.Bucket })
    $ghCols  = @($columns | Where-Object { $null -ne $_.Gh })
    Write-MatrixTable 'Serious Awards' @($gitRows | Where-Object { $_.Award.Cat -eq 'Serious' }) $gitCols
    Write-MatrixTable 'Funny Awards'   @($gitRows | Where-Object { $_.Award.Cat -eq 'Funny' }) $gitCols
    if ($ghRows.Count) { Write-MatrixTable 'GitHub Awards' $ghRows $ghCols }
}

function Write-RankingList($r) {
    for ($i = 0; $i -lt $r.Count; $i++) {
        $report.Add("- $(Get-Medal $i): **$($r[$i].Name)** ($($r[$i].Text))")
    }
    $report.Add('')
}

function Write-SectionsLayout {
    $report.Add('Every award is calculated separately for each repo' +
                $(if ($NoCombined) { '.' } else { ', followed by the standings for all repos combined.' }))
    $report.Add('')
    foreach ($col in $columns) {
        $report.Add($(if ($col.Combined) { '## All Repos Combined' } else { "## $($col.Label)" }))
        $report.Add("_$($col.Note)_"); $report.Add('')
        if ($col.Bucket) {
            $section = $null
            foreach ($row in $gitRows) {
                $r = $row.Cells[$col.Label]
                if (-not $r -or $r.Count -eq 0) { continue }
                if ($row.Award.Cat -ne $section) {
                    $section = $row.Award.Cat
                    $report.Add("### $section Awards"); $report.Add('')
                }
                $report.Add("#### $($row.Award.Title)"); $report.Add("_$($row.Award.Blurb)_")
                Write-RankingList $r
            }
        }
        if ($null -ne $col.Gh) {
            $report.Add($(if ($col.Slug) { "### GitHub Awards ($($col.Slug))" } else { '### GitHub Awards' })); $report.Add('')
            foreach ($row in $ghRows) {
                if ($row.Award.CombinedOnly -and -not $col.Combined) { continue }
                $r = $row.Cells[$col.Label]
                $report.Add("#### $($row.Award.Title)"); $report.Add("_$($row.Award.Blurb)_")
                if (-not $r -or $r.Count -eq 0) { $report.Add('- (no qualifying data)'); $report.Add(''); continue }
                Write-RankingList $r
            }
        }
    }
    foreach ($label in $emptyRepos) {
        $report.Add("## $label"); $report.Add('_No commits in this period._'); $report.Add('')
    }
}

if ($Layout -eq 'Sections') { Write-SectionsLayout } else { Write-MatrixLayout }

# ---- summary
$report.Add('## Award Winners at a Glance'); $report.Add('')
$report.Add('_Each #1 finish, with the repo it came from in parentheses._'); $report.Add('')
foreach ($name in ($winners.Keys | Sort-Object { -(Get-WinCount $winners[$_]) }, { $_ })) {
    $w = $winners[$name]
    $n = Get-WinCount $w
    $parts = foreach ($title in $w.Keys) { "$title ($($w[$title] -join ', '))" }
    $report.Add("- **$name** ($n win$(if ($n -ne 1) { 's' })): $($parts -join '; ')")
}
$leftOut = @($all.Stats.Keys | Where-Object { -not $winners.ContainsKey($_) -and $all.Stats[$_].commits -gt 0 } | Sort-Object)
if ($leftOut.Count) {
    $report.Add(''); $report.Add('## No #1 finish yet (consider a runner-up award!)'); $report.Add('')
    foreach ($name in $leftOut) { $report.Add("- $name ($(Format-Num $all.Stats[$name].commits) commits)") }
}

$text = $report -join "`n"
$text
if ($Out) {
    Set-Content -Path $Out -Value $text -Encoding utf8
    Write-Host "`nSaved to $Out"
}
