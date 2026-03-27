# 设置捕获机制，防止任何错误导致闪退
try {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

    Write-Host "--- SEO 交付物打包助手 Pro (增强稳定版) ---" -ForegroundColor Yellow

    # 1. 选择洁净文档目录
    $FolderBrowser.Description = "第1步：请选择【Cleaned_Output】文件夹"
    if ($FolderBrowser.ShowDialog() -ne "OK") { throw "用户取消了选择：Cleaned_Output" }
    $CleanedDir = $FolderBrowser.SelectedPath

    # 2. 选择原始资源目录
    $FolderBrowser.Description = "第2步：请选择【原始资源】文件夹"
    if ($FolderBrowser.ShowDialog() -ne "OK") { throw "用户取消了选择：原始资源" }
    $OriginalDir = $FolderBrowser.SelectedPath

    # 3. 设置截图过滤选项
    Write-Host "`n[选项确认]" -ForegroundColor Cyan
    $ExcludeScreenshots = Read-Host "是否排除文件名中包含 'screenshot' 的图片？(Y/N, 默认Y)"
    if ($ExcludeScreenshots -eq "" -or $ExcludeScreenshots -eq "y") { $ExcludeScreenshots = $true } else { $ExcludeScreenshots = $false }

    # 4. 准备目标路径 (桌面)
    $Timestamp = Get-Date -Format "MMdd_HHmm"
    $FinalDelivery = Join-Path ([Environment]::GetFolderPath("Desktop")) "Delivery_$Timestamp"

    # 5. 预扫描
    Write-Host "`n[正在扫描文件...]" -ForegroundColor Cyan
    if (-not (Test-Path $CleanedDir)) { throw "找不到路径: $CleanedDir" }
    
    $DocFiles = Get-ChildItem -Path $CleanedDir -Recurse -File -Filter *.docx
    $ImgQuery = Get-ChildItem -Path $OriginalDir -Recurse -File -Include *.webp, *.png, *.jpg, *.jpeg
    
    if ($ExcludeScreenshots) {
        $ImgFiles = $ImgQuery | Where-Object { $_.Name -notmatch "screenshot" }
    } else {
        $ImgFiles = $ImgQuery
    }

    Write-Host "------------------------------------"
    Write-Host "待处理文档：$($DocFiles.Count) 个"
    Write-Host "待处理图片：$($ImgFiles.Count) 张"
    Write-Host "输出目标：$FinalDelivery"
    Write-Host "------------------------------------"
    
    Write-Host "确认无误请按回车开始，或点击右上角关闭..." -ForegroundColor Gray
    [void][System.Console]::ReadLine()

    # 6. 执行复制
    New-Item -ItemType Directory -Path $FinalDelivery -Force | Out-Null
    Copy-Item -Path "$CleanedDir\*" -Destination $FinalDelivery -Recurse -Force

    foreach ($file in $ImgFiles) {
        $RelativePath = $file.FullName.Substring((Get-Item $OriginalDir).FullName.Length)
        $TargetFilePath = Join-Path $FinalDelivery $RelativePath
        $TargetDir = Split-Path $TargetFilePath
        
        if (-not (Test-Path $TargetDir)) { New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null }
        Copy-Item -Path $file.FullName -Destination $TargetFilePath -Force
    }

    Write-Host "`n[打包成功!] 交付文件夹已在桌面生成。" -ForegroundColor Green
}
catch {
    Write-Host "`n[脚本出错!]" -ForegroundColor Red -BackgroundColor Black
    Write-Host "错误原因: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "`n请截图此窗口反馈，按任意键退出..."
    [void][System.Console]::ReadKey()
}
finally {
    # 确保无论如何最后都有个停顿
    Write-Host "`n--- 运行结束 ---"
}