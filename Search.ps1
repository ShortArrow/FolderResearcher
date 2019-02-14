# 複数フォルダを指定できる設計になっています(配列)。

# System.Windows.Formsを有効化
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# 基本設定
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = "客先名フォルダを選択してください。" 
$fbd.SelectedPath = "\\192.168.0.170\supersub\図面Tiffデータ"
$TargetFolders = @("出図連絡書", "変更リスト", "仕様書")

# ダイアログを表示する
$target = $fbd.ShowDialog() | Out-Null
# 選択をキャンセルした場合はNULLを返す
if ( $target -eq [System.Windows.Forms.DialogResult]::Cancel) {
    $targetPath = $null
}
else {
    $targetPath = $fbd.SelectedPath 
}
Write-Host $targetPath
$FoldersConfigPath = $targetPath
# $FoldersConfigPath = "\\192.168.0.170\supersub\図面Tiffデータ\TOYOTA"
$DIRS = (Get-ChildItem $FoldersConfigPath -Directory) -as [string[]]

# メインの処理
# 実行中のパス取得/移動
$path = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $path
$FoldersConfigPathSplit = $FoldersConfigPath.Split("\")
$fileName = $path + "\log\$($FoldersConfigPathSplit[$FoldersConfigPathSplit.Length-1]).csv"
$file = New-Object System.IO.StreamWriter($fileName, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$file.Write("工番,図面数")
foreach ($item in $TargetFolders) {
    $file.Write(",$($item)フォルダ,$($item)数")
}
$file.WriteLine("")
$Counter = 1
foreach ($DIR in $DIRS) {
    # $finderPath = ("FileSystem::$DIR") # dir $DIR/hoge.xlsx にすれば、hoge.xlsxファイルだけに絞れます
    $finderPath = $FoldersConfigPath + "\" + $DIR
    $Folders = (Get-ChildItem $finderPath -Directory -Depth 0) -as [string[]]
    $Files = (Get-ChildItem $finderPath -File -Depth 0 -Name) -as [string[]]
    # 図面
    $Drawings = ($Files -match ".+\.tif?") -as [string[]]
    $DrawingsCount = $Drawings.Length
    $file.Write("=""" + $DIR + """," + $DrawingsCount)
    foreach ($item in $TargetFolders) {
        $FolderCounter = $(($Folders -match "$item.*") -as [string[]]).Length
        if (Test-Path ($finderPath + "\" + $item)) {
            $FileCounter = $((Get-ChildItem ($finderPath + "\" + $item) -File -Depth 0 -Name) -as [string[]]).Length
        }
        else {
            $FileCounter = 0
        }
        $file.Write("," + $FolderCounter + "," + $FileCounter)
    }
    $file.WriteLine("")
    $activity = "図面TIFF内部　オートチェッカー"
    $status = "CSV書き込み中"
    $ProgressRate = [Math]::Round(($Counter / $DIRS.Length) * 100, 2, [MidpointRounding]::AwayFromZero)
    Write-Progress $activity $status -PercentComplete $ProgressRate -CurrentOperation "$ProgressRate%完了"
    Start-Sleep -Milliseconds 10
    $Counter++
    if ($DIR -eq "00020") {
        # break
    }
}
$file.Close()


# アセンブリの読み込み
Add-Type -Assembly System.Windows.Forms | Out-Null
#結果表示
[void][System.Windows.Forms.MessageBox]::Show("Finish Research", "GF企画室")
Write-Host
Write-Host "Finished Please Input Any Key"
exit
