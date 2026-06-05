param([string]$src, [string]$deploy)

Write-Host '   [1/4] Reading table...'
try {
    $ext = [IO.Path]::GetExtension($src).ToLower()
    if ($ext -eq '.csv') {
        $lines = Get-Content $src -Encoding UTF8
        if ($lines.Count -lt 2) { throw 'CSV too short' }
        $headers = $lines[0] -split "`t|," | ForEach-Object { $_.Trim().Trim('"') }
        $rows = @()
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $vals = $lines[$i] -split "`t|," | ForEach-Object { $_.Trim().Trim('"') }
            $obj = @{}
            for ($j = 0; $j -lt [Math]::Min($headers.Count, $vals.Count); $j++) { $obj[$headers[$j]] = $vals[$j] }
            $rows += $obj
        }
    } else {
        # Try WPS first, then Excel
        $app = $null
        $progId = ''
        foreach ($pname in @('Ket.Application', 'Excel.Application')) {
            try {
                $app = New-Object -ComObject $pname
                $progId = $pname
                break
            } catch { }
        }
        if (-not $app) { throw 'No WPS or Excel found on this PC' }
        Write-Host ('   Using: ' + $progId)
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $wb = $app.Workbooks.Open($src)
        $ws = $wb.Worksheets.Item(1)
        $used = $ws.UsedRange
        $data = $used.Value2
        if ($data.Count -lt 2) { $wb.Close($false); $app.Quit(); throw 'Table too short' }
        $headers = @()
        $rowCount = $data.GetLength(0)
        $colCount = $data.GetLength(1)
        for ($c = 1; $c -le $colCount; $c++) {
            $h = $data[1, $c]
            if ($h) { $headers += $h.ToString().Trim() } else { $headers += '' }
        }
        $rows = @()
        for ($r = 2; $r -le $rowCount; $r++) {
            $obj = @{}
            for ($c = 1; $c -le $colCount; $c++) {
                $key = $headers[$c - 1]
                $val = $data[$r, $c]
                if ($val -is [double] -and $val -gt 30000 -and $val -lt 80000 -and ($key -like '*日期*' -or $key -like '*时间*')) {
                    $val = [DateTime]::FromOADate($val).ToString('yyyy-MM-dd')
                }
                if ($null -eq $val) { $val = '' }
                $obj[$key] = $val.ToString().Trim()
            }
            $rows += $obj
        }
        $wb.Close($false)
        $app.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }

    Write-Host ('   Read ' + $rows.Count + ' records')
    Write-Host '   [2/4] Convert to JSON...'
    # 附加 _updateTime 元字段，供网页显示真实数据更新时间
    $updateTimeStr = Get-Date -Format 'yyyy/MM/dd HH:mm'
    $metaRow = @{ '_updateTime' = $updateTimeStr }
    $payload = @{ meta = $metaRow; data = $rows }
    $json = ConvertTo-Json -InputObject $payload -Depth 10 -Compress
    if (-not (Test-Path $deploy)) { New-Item -ItemType Directory -Path $deploy -Force | Out-Null }
    $jsonPath = Join-Path $deploy 'data.json'
    [IO.File]::WriteAllText($jsonPath, $json, [Text.Encoding]::UTF8)
    $sizeKB = [Math]::Round((Get-Item $jsonPath).Length / 1024, 1)
    Write-Host ('   JSON size: ' + $sizeKB + ' KB')

    Write-Host '   [3/4] Git commit...'
    # Locate git.exe (may not be in PowerShell PATH)
    $gitExe = $null
    foreach ($gpath in @(
        "$env:USERPROFILE\.workbuddy\vendor\PortableGit\cmd\git.exe",
        "$env:USERPROFILE\.workbuddy\vendor\PortableGit\bin\git.exe",
        'C:\Program Files\Git\cmd\git.exe',
        'C:\Program Files\Git\bin\git.exe',
        "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe",
        "$env:LOCALAPPDATA\Programs\Git\bin\git.exe"
    )) { if (Test-Path $gpath) { $gitExe = $gpath; break } }
    if (-not $gitExe) { $gitExe = 'git' }  # fallback to PATH
    Write-Host ('   Git: ' + $gitExe)
    Push-Location $deploy
    $dateStr = Get-Date -Format 'yyyy-MM-dd'
    & $gitExe add data.json
    $addOk = ($LASTEXITCODE -eq 0)
    & $gitExe commit -m ('update data ' + $dateStr)
    $commitOk = ($LASTEXITCODE -eq 0)

    # 推送到 GitHub（带重试：网络问题重试，冲突则自动 rebase）
    function Push-WithRetry {
        param($gitPath)
        $maxTries = 3
        for ($i = 1; $i -le $maxTries; $i++) {
            if ($i -gt 1) {
                Write-Host ('   Retry ' + ($i) + '/' + $maxTries + ' after 5s...')
                Start-Sleep -Seconds 5
            }
            $output = & $gitPath push origin main 2>&1
            if ($LASTEXITCODE -eq 0) { return $true }
            # 如果远程有新提交导致冲突，先 rebase 再重试
            if ($output -match 'rejected') {
                Write-Host '   Remote has newer commits, pulling & rebasing...'
                & $gitPath pull --rebase origin main 2>&1 | Out-Null
                $output = & $gitPath push origin main 2>&1
                if ($LASTEXITCODE -eq 0) { return $true }
            }
        }
        return $false
    }

    if ($commitOk) {
        Write-Host '   [4/4] Push to GitHub...'
        $pushOk = Push-WithRetry $gitExe
        Pop-Location

        if ($pushOk) {
            Write-Host ''
            Write-Host '   ========================================'
            Write-Host '   [SUCCESS] New data pushed to GitHub!'
            Write-Host '   Refresh: https://zhongshanms.github.io/zhifa-dashboard/'
            Write-Host '   ========================================'
        } else {
            Write-Host ''
            Write-Host '   ========================================'
            Write-Host '   [FAILED] Push failed after 3 retries!'
            Write-Host '   Data committed locally. Re-run script when network is back.'
            Write-Host '   ========================================'
        }
    } else {
        Pop-Location
        # commit 跳过，但可能有之前没推成功的积压 commit
        $ahead = & $gitExe rev-list --count origin/main..HEAD 2>$null
        if ($LASTEXITCODE -eq 0 -and [int]$ahead -gt 0) {
            Write-Host '   [PUSH] ' + $ahead + ' pending commit(s), pushing...'
            Push-Location $deploy
            $pushOk = Push-WithRetry $gitExe
            Pop-Location
            if ($pushOk) {
                Write-Host ''
                Write-Host '   ========================================'
                Write-Host '   [SUCCESS] Pending commits pushed!'
                Write-Host '   Refresh: https://zhongshanms.github.io/zhifa-dashboard/'
                Write-Host '   ========================================'
            } else {
                Write-Host ''
                Write-Host '   ========================================'
                Write-Host '   [FAILED] Push failed after 3 retries!'
                Write-Host '   Check network and try again.'
                Write-Host '   ========================================'
            }
        } else {
            Write-Host ''
            Write-Host '   ========================================'
            Write-Host '   [INFO] Data unchanged - already up to date.'
            Write-Host '   ' + $rows.Count + ' records, ' + $sizeKB + ' KB'
            Write-Host '   ========================================'
        }
    }
} catch {
    Write-Host ''
    Write-Host '   ========================================'
    Write-Host ('   [ERROR] ' + $_.Exception.Message)
    Write-Host '   ========================================'
    Write-Host ''
    Write-Host '   If Ludun encryption issue:'
    Write-Host '   1. Open file in WPS/Excel'
    Write-Host '   2. Save As CSV'
    Write-Host '   3. Drag CSV here'
}
