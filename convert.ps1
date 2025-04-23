# usage
# .\this.ps1 -PPSXFilePath "C:\path\to\your\file.ppsx"

param(
    [Parameter(Mandatory=$true)]
    [string]$PPSXFilePath
)

# PowerPointアプリケーションのオブジェクトを作成
$PowerPoint = New-Object -ComObject PowerPoint.Application

try {
    # 指定されたPPSXファイルが存在するか確認
    if (-not (Test-Path -Path $PPSXFilePath -PathType Leaf)) {
        Write-Error "指定されたファイルが見つかりません: $PPSXFilePath"
        exit 1
    }

    # プレゼンテーションを開く（読み取り専用、非表示）
    $Presentation = $PowerPoint.Presentations.Open($PPSXFilePath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    # 新しいファイル名を作成（拡張子をpptxに変更）
    $PPTXFilePath = [System.IO.Path]::ChangeExtension($PPSXFilePath, ".pptx")

    # PPTX形式で保存
    $Presentation.SaveAs($PPTXFilePath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsDefault)

    # プレゼンテーションを閉じる
    $Presentation.Close()

    Write-Host "Successfully converted: $($PPSXFilePath) to $($PPTXFilePath)"
}
catch {
    Write-Error "Failed to convert: $($PPSXFilePath) - $($_.Exception.Message)"
    exit 1
}
finally {
    # PowerPointアプリケーションを終了 (tryまたはcatchブロックの後に必ず実行)
    if ($PowerPoint) {
        $PowerPoint.Quit()
    }
}