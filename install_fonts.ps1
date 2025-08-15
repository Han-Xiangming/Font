# =================================================================
# PowerShell 字体批量安装工具 (终极版 v24 - The Definitive Edition)
#
# 最终修正:
#   - 1. [作用域修正] 彻底解决了在 ForEach-Object -Parallel 中因非法使用
#      $using: 表达式而导致的致命解析器错误。
#   - 2. [代码模式优化] 采用"先引用、再操作"的最佳实践，确保并行线程的
#      语法健壮性和稳定性。
#   - 本版本是整个优化旅程的最终稳定成果。
# =================================================================

#region Function Definitions
function Show-Header {
    Clear-Host
    $title = "PowerShell 字体批量安装工具 (v24 - The Definitive Edition)"
    $line = "=" * ($title.Length + 4); Write-Host $line -F Green; Write-Host "  $title  " -F White; Write-Host $line -F Green
    $Host.UI.RawUI.WindowTitle = $title
}
function Write-ReportLine { param([string]$Label, [string]$Value, [string]$ValueColor = 'White', [int]$LabelWidth = 18) $paddedLabel = $Label.PadRight($LabelWidth); $paddedValue = $Value.ToString().PadLeft(8); Write-Host -NoNewline ("  {0} : " -f $paddedLabel) -F White; Write-Host $paddedValue -F $ValueColor }
function Confirm-Choice { param( [string]$Prompt, [string]$DefaultChoice = 'Y' ) $promptText = if($DefaultChoice -eq 'Y') { "$Prompt [Y/N] (默认为: Y)" } else { "$Prompt [Y/N] (默认为: N)" }; while ($true) { Write-Host -NoNewline $promptText; $keyInfo = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); $char = $keyInfo.Character.ToString().ToUpper(); if ($keyInfo.VirtualKeyCode -eq 13) { Write-Host " -> $DefaultChoice"; return $DefaultChoice -eq 'Y' }; if ("Y", "N" -contains $char) { Write-Host " -> $char"; return $char -eq 'Y' } } }
function Get-UserConfiguration { param($CachedFontCount) Write-Host "`n欢迎使用！本工具将引导您完成字体安装的配置。`n" -F White; Write-Host "后台已为您预先缓存了 $CachedFontCount 种系统字体信息...`n" -F Gray; $targetDir = ""; do { Write-Host "--- [步骤 1/3] 请指定字体文件所在的目录 ---" -F Cyan; $defaultPath = (Get-Location).Path; $inputPath = Read-Host "输入路径 (默认为: $defaultPath)"; if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = $defaultPath }; $targetDir = Resolve-Path -Path $inputPath -ErrorAction SilentlyContinue; if (-not $targetDir -or -not (Test-Path $targetDir -PathType Container)) { Write-Host "`n[错误] 目录 ""$inputPath"" 无效或不存在，请重新输入。`n" -F Red; $targetDir = $null } } until ($targetDir); Write-Host "`n--- [步骤 2/3] 请选择安装模式 ---" -F Cyan; $overwriteFonts = Confirm-Choice -Prompt "是否覆盖已存在的同名字体?" -DefaultChoice 'N'; Write-Host "`n--- [步骤 3/3] 请确认您的配置 ---" -F Cyan; $line = "-" * 50; Write-Host $line; Write-Host ("{0,-12} : {1}" -f "目标目录", $targetDir); Write-Host ("{0,-12} : {1}" -f "安装模式", $(if ($overwriteFonts) { "覆盖安装" } else { "标准安装" })) -F $(if ($overwriteFonts) { "Yellow" } else { "White" }); Write-Host $line; if (-not (Confirm-Choice -Prompt "`n确认开始执行吗?" -DefaultChoice 'Y')) { return $null }; return [PSCustomObject]@{ TargetDirectory = $targetDir; OverwriteMode = $overwriteFonts } }
function Show-FinalReport { param($Params) $modeText = if ($Params.Config.OverwriteMode) { "覆盖安装" } else { "标准安装" }; $filesPerSec = if ($Params.ProcessingTime.TotalSeconds -gt 0) { [math]::Round($Params.TotalFiles / $Params.ProcessingTime.TotalSeconds, 2) } else { 0 }; Write-Host "`n`n====================== [ 指挥中心最终报告 ] ======================" -F Green; Write-Host "`n  --- [ 性能分析 ] ---" -F Cyan; Write-ReportLine -Label "总耗时" -Value $Params.TotalTime.ToString("g"); Write-ReportLine -Label "  - 缓存耗时" -Value "$($Params.CacheTime.TotalMilliseconds) ms"; Write-ReportLine -Label "  - 处理耗时" -Value $Params.ProcessingTime.ToString("g"); Write-ReportLine -Label "  - 注册耗时" -Value "$($Params.RegisterTime.TotalMilliseconds) ms"; Write-ReportLine -Label "平均处理速度" -Value "$filesPerSec 文件/秒"; Write-Host "`n  --- [ 数据详单与可视化 ] ---" -F Cyan; Write-ReportLine -Label "共发现字体文件" -Value $Params.TotalFiles; Write-ReportLine -Label "立即安装 (新字体)" -Value $Params.InstalledFiles -ValueColor Green; Write-ReportLine -Label "计划在重启后更新" -Value $Params.RebootNeededFiles -ValueColor Cyan; Write-ReportLine -Label "跳过 (已存在)" -Value $Params.SkippedFiles -ValueColor Yellow; if ($Params.FailedFiles -gt 0) { Write-ReportLine -Label "安装/更新失败" -Value $Params.FailedFiles -ValueColor Red }; $barWidth = 50; Write-Host; if ($Params.TotalFiles -gt 0) { $instBlocks=[math]::Round($barWidth*$Params.InstalledFiles/$Params.TotalFiles); $rebootBlocks=[math]::Round($barWidth*$Params.RebootNeededFiles/$Params.TotalFiles); $failBlocks=[math]::Round($barWidth*$Params.FailedFiles/$Params.TotalFiles); $skipBlocks=$barWidth-$instBlocks-$rebootBlocks-$failBlocks; Write-Host -NoNewline "  ["; Write-Host -NoNewline ('█'*$instBlocks) -F Green; Write-Host -NoNewline ('█'*$rebootBlocks) -F Cyan; Write-Host -NoNewline ('█'*$failBlocks) -F Red; Write-Host -NoNewline ('█'*$skipBlocks) -F Yellow; Write-Host "]"; Write-Host -NoNewline "  图例: "; Write-Host -NoNewline "█" -F Green; Write-Host -NoNewline " 新装  "; Write-Host -NoNewline "█" -F Cyan; Write-Host -NoNewline " 更新  "; if ($Params.FailedFiles -gt 0) { Write-Host -NoNewline "█" -F Red; Write-Host -NoNewline " 失败  " }; Write-Host -NoNewline "█" -F Yellow; Write-Host " 跳过" }; if ($Params.ErrorLog.Count -gt 0) { $logHeaderText = " 问题日志 (共 $($Params.ErrorLog.Keys.Count) 种错误) "; Write-Host "`n  ---$logHeaderText---" -F Yellow; $Params.ErrorLog.GetEnumerator() | Sort-Object Value -Desc | % { Write-Host -NoNewline ("   [发生 " + ($_.Value).ToString().PadLeft(3) + " 次] ") -F Red; Write-Host $_.Key -F Gray } }; Write-Host "`n  --- [ 行动指令 ] ---" -F Cyan; if ($Params.RebootNeededFiles -gt 0) { $line = "*" * 62; Write-Host "`n  $line" -F Red; Write-Host ("  * {0, -58} *" -f "重要: 需要重启计算机以完成所有字体更新。") -F White; Write-Host "  $line" -F Red } elseif ($Params.InstalledFiles > 0) { Write-Host "  建议: 注销后重新登录，以确保所有程序都能加载新字体。" -F Yellow } else { Write-Host "  操作已完成，未发生任何需要用户干预的更改。" }; Write-Host "`n==================================================================" -F Green }

function Main {
    Show-Header
    # [兼容性修正] 严格遵守 Here-String 语法规范
    if (-not ("NativeMethods" -as [type])) {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public static class NativeMethods {
            [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
            public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
            public const int MOVEFILE_DELAY_UNTIL_REBOOT = 0x4;
        }
"@
    }
    
    $ts = "[{0:HH:mm:ss}]" -f (Get-Date); Write-Host "`n$ts [ 准备阶段 ]" -ForegroundColor Cyan
    $cacheStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "正在使用 .NET API 高速缓存系统字体..." -NoNewline
    $installedFontMap = @{}
    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Default').OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
        foreach($valueName in $regKey.GetValueNames()){
            $fileName = $regKey.GetValue($valueName)
            if ($fileName -and -not $installedFontMap.ContainsKey($fileName)) { $installedFontMap.Add($fileName, $valueName) }
        }
        $regKey.Close()
    } catch {}; $cacheStopwatch.Stop()
    Write-Host " 完成！缓存了 $($installedFontMap.Count) 个字体 (耗时: $($cacheStopwatch.Elapsed.TotalMilliseconds)ms)。" -ForegroundColor Green
    
    $config = Get-UserConfiguration -CachedFontCount $installedFontMap.Count
    if (-not $config) { Write-Host "`n操作已由用户取消。"; Write-Host "`n请按任意键退出..."; $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null; return }

    $masterStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $fontExtensions = @("*.ttf", "*.otf", "*.ttc", "*.otc"); $regImportPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"; $regImportFile = Join-Path $env:TEMP "font_install_$([System.Guid]::NewGuid()).reg"; $fontsTargetDir = Join-Path $env:windir "Fonts"
    $regKeyAllUsers = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    
    $ts = "[{0:HH:mm:ss}]" -f (Get-Date); Write-Host "`n$ts [ 执行阶段 ]" -ForegroundColor Cyan
    
    Write-Host "正在扫描并收集文件列表..." -NoNewline
    $fontFiles = Get-ChildItem -Path $config.TargetDirectory -Include $fontExtensions -Recurse -File -ErrorAction SilentlyContinue
    $totalFiles = $fontFiles.Count
    Write-Host " 完成！发现 $totalFiles 个文件。" -ForegroundColor Green
    if ($totalFiles -eq 0) { $masterStopwatch.Stop(); Write-Host "`n请按任意键退出..."; $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null; return }

    $processingStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $ts = "[{0:HH:mm:ss}]" -f (Get-Date); Write-Host "$ts 正在启动并行分析引擎..."
    $throttleLimit = [Environment]::ProcessorCount
    $actionList = $fontFiles | ForEach-Object -ThrottleLimit $throttleLimit -Parallel {
        # [作用域修正] 先将 using 变量存入本地变量
        $localMap = $using:installedFontMap
        $localOverwrite = $using:config.OverwriteMode
        
        $font = $_
        $isInstalled = $localMap.ContainsKey($font.Name)
        
        if (-not $isInstalled) { [PSCustomObject]@{ Action='Install'; Font=$font } }
        elseif ($isInstalled -and $localOverwrite) { [PSCustomObject]@{ Action='Overwrite'; Font=$font } }
        else { [PSCustomObject]@{ Action='Skip'; Font=$font } }
    }
    $processingStopwatch.Stop()
    $ts = "[{0:HH:mm:ss}]" -f (Get-Date); Write-Host "$ts 并行分析完成 (耗时: $($processingStopwatch.Elapsed.ToString("g")))。正在执行 I/O 操作..."

    $skippedFiles = 0; $installedFiles = 0; $rebootNeededFiles = 0; $failedFiles = 0; $errorLog = @{}; $fontsToRegister = 0
    $regBuilder = New-Object System.Text.StringBuilder; $regBuilder.AppendLine("Windows Registry Editor Version 5.00`n[$regImportPath]") | Out-Null
    foreach ($item in $actionList) {
        switch ($item.Action) {
            'Install' {
                try { Copy-Item -Path $item.Font.FullName -Destination $fontsTargetDir -Force -ErrorAction Stop; $installedFiles++; $fontsToRegister++; $regBuilder.AppendLine("""$($item.Font.BaseName) (TrueType)""=""$($item.Font.Name)""") | Out-Null } 
                catch { if ($errorLog.ContainsKey($_.Exception.Message)) { $errorLog[$_.Exception.Message]++ } else { $errorLog[$_.Exception.Message] = 1 }; $failedFiles++ }
            }
            'Overwrite' {
                $operationSucceeded = $false
                try { if ($installedFontMap.ContainsKey($item.Font.Name)) { Remove-ItemProperty -Path $regKeyAllUsers -Name $installedFontMap[$item.Font.Name] -Force -ErrorAction Stop }
                    $tempFontPath = "$(Join-Path $fontsTargetDir $item.Font.Name).$([System.Guid]::NewGuid()).tmp"; Copy-Item -Path $item.Font.FullName -Destination $tempFontPath -Force -ErrorAction Stop
                    if ([NativeMethods]::MoveFileEx($tempFontPath, (Join-Path $fontsTargetDir $item.Font.Name), [NativeMethods]::MOVEFILE_DELAY_UNTIL_REBOOT)) { $operationSucceeded = $true } 
                    else { $errMsg = "为 `"$($item.Font.Name)`" 注册重启替换任务失败。"; if ($errorLog.ContainsKey($errMsg)) { $errorLog[$errMsg]++ } else { $errorLog[$errMsg] = 1 } }
                } catch { if ($errorLog.ContainsKey($_.Exception.Message)) { $errorLog[$_.Exception.Message]++ } else { $errorLog[$_.Exception.Message] = 1 } }
                if ($operationSucceeded) { $rebootNeededFiles++; $fontsToRegister++; $regBuilder.AppendLine("""$($item.Font.BaseName) (TrueType)""=""$($item.Font.Name)""") | Out-Null } else { $failedFiles++ }
            }
            'Skip' { $skippedFiles++ }
        }
    }
    
    $registerStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    if ($fontsToRegister -gt 0) { Write-Host "`n正在执行批量注册..."; Set-Content -Path $regImportFile -Value $regBuilder.ToString() -Encoding Ascii; Start-Process "reg.exe" -ArgumentList "import `"$regImportFile`"" -Wait -WindowStyle Hidden; Remove-Item $regImportFile -Force; Write-Host "批量注册成功！" -ForegroundColor Green }
    $registerStopwatch.Stop(); $masterStopwatch.Stop()
    
    $reportParams = @{ Config=$config; TotalTime=$masterStopwatch.Elapsed; CacheTime=$cacheStopwatch.Elapsed; ProcessingTime=$processingStopwatch.Elapsed; RegisterTime=$registerStopwatch.Elapsed; TotalFiles=$totalFiles; SkippedFiles=$skippedFiles; InstalledFiles=$installedFiles; RebootNeededFiles=$rebootNeededFiles; FailedFiles=$failedFiles; ErrorLog=$errorlog }
    Show-FinalReport -Params $reportParams
    
    Write-Host "`n操作完成，请按任意键退出..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
}
#endregion

# --- 启动程序 ---
Main