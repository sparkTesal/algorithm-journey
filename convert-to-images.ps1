# 添加System.Drawing命名空间以使用图形功能
Add-Type -AssemblyName System.Drawing

function Merge-ImagesToLongImage($imageFiles, $outputImagePath) {
    $totalHeight = 0
    $maxWidth = 0
    $images = @()

    # 计算总高度和最大宽度
    foreach ($imageFile in $imageFiles) {
        $img = [System.Drawing.Image]::FromFile($imageFile)
        $totalHeight += $img.Height
        $maxWidth = [math]::Max($maxWidth, $img.Width)
        $images += $img
    }

    # 创建一个足够大的bitmap来容纳所有图片
    $finalImage = New-Object System.Drawing.Bitmap $maxWidth, $totalHeight
    $graphics = [System.Drawing.Graphics]::FromImage($finalImage)
    $graphics.Clear([System.Drawing.Color]::White)

    $currentY = 0
    # 将每个图片绘制到最终图片上
    foreach ($img in $images) {
        $graphics.DrawImage($img, 0, $currentY)
        $currentY += $img.Height
        $img.Dispose()
    }

    # 保存最终的图片
    $finalImage.Save($outputImagePath)

    # 释放资源
    $graphics.Dispose()
    $finalImage.Dispose()
}
# 获取脚本文件所在的目录
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
# 设置工作目录为包含PPT文件的目录
$folderPath = Join-Path -Path $scriptPath -ChildPath "ppt"
# 输出图片的目录
$outputFolderPath = Join-Path -Path $scriptPath -ChildPath "image"

# 如果输出目录不存在，则创建
if (-not (Test-Path -Path $outputFolderPath)) {
    New-Item -ItemType Directory -Force -Path $outputFolderPath
}

# 创建一个新的PowerPoint应用实例
$pptApplication = New-Object -ComObject PowerPoint.Application

# 获取所有PPT文件
$pptFiles = Get-ChildItem -Path $folderPath -Filter *.pptx

foreach ($file in $pptFiles) {
    $imageFiles = New-Object System.Collections.ArrayList

    # 打开PPT文件
    $presentation = $pptApplication.Presentations.Open($file.FullName)

    # 将每张幻灯片保存为图片文件
    $tempFolderPath = [System.IO.Path]::GetTempPath()
    foreach ($slide in $presentation.Slides) {
        $tempImageFilePath = [System.IO.Path]::Combine($tempFolderPath, [GUID]::NewGuid().ToString() + ".jpg")
        $slide.Export($tempImageFilePath, "JPG")
        $imageFiles.Add($tempImageFilePath)
    }

    # 关闭PPT文件
    $presentation.Close()

    # 合并当前PPT导出的所有图片到一个长图
    $outputImagePath = Join-Path -Path $outputFolderPath -ChildPath ($file.BaseName + ".jpg")
    Merge-ImagesToLongImage $imageFiles $outputImagePath

    # 删除临时图片文件
    $imageFiles | ForEach-Object { Remove-Item $_ }
}

# 释放COM对象，退出PowerPoint
$pptApplication.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApplication)