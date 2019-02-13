# 複数フォルダを指定できる設計になっています(配列)。

# System.Windows.Formsを有効化
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# 必要な情報を設定
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = "対象ディレクトリを選択してください。" 
$fbd.SelectedPath = "\\192.168.0.170\supersub\図面Tiffデータ"

# ダイアログを表示する
# $target = $fbd.ShowDialog() | Out-Null

# 選択をキャンセルした場合はNULLを返す
if ( $target -eq [System.Windows.Forms.DialogResult]::Cancel) {
    $targetPath = $null
}
else {
    $targetPath = $fbd.SelectedPath 
}

Write-Host $targetPath

# $FoldersConfigPath = $targetPath
$FoldersConfigPath = "\\192.168.0.170\supersub\図面Tiffデータ\TOYOTA"
$DIRS = (Get-ChildItem $FoldersConfigPath -Directory) -as [string[]]
# メインの処理
$errorMessage = ""
$resultMessage = ""
# Write-Host $DIRS.Contains("00007").ToString()
# Write-Host $DIRS.Contains("00000").ToString()

# 実行中のパス取得/移動
$path = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $path
$fileName = $path + "\Log.csv"
$file = New-Object System.IO.StreamWriter($fileName, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$file.WriteLine("""工番"",""出連フォルダ"",""変更リストフォルダ"",""仕様書フォルダ"",""図面数""")
$Counter = 1
foreach ($DIR in $DIRS) {
    # $finderPath = ("FileSystem::$DIR") # dir $DIR/hoge.xlsx にすれば、hoge.xlsxファイルだけに絞れます
    $finderPath = $FoldersConfigPath + "\" + $DIR
    $Folders = (Get-ChildItem $finderPath -Directory -Depth 0) -as [string[]]
    $Files = (Get-ChildItem $finderPath -File -Depth 0 -Name) -as [string[]]
    # 出図連絡書
    $EjectDrawingSheetFolders = ($Folders -match "出図連絡書.*") -as [string[]]
    $EjectDrawingSheetFoldersCount = $EjectDrawingSheetFolders.Length
    # 変更リスト
    $ChangeListFolders = ($Folders -match "変更リスト.*") -as [string[]]
    $ChangeListFoldersCount = $ChangeListFolders.Length
    # 仕様書
    $SpecSheetFolders = ($Folders -match "仕様書.*") -as [string[]]
    $SpecSheetFoldersCount = $SpecSheetFolders.Length
    # 図面
    $Drawings = ($Files -match ".+\.tif?") -as [string[]]
    $DrawingsCount = $Drawings.Length
    

    $file.WriteLine("""" + $DIR + """,""" + $EjectDrawingSheetFoldersCount + """,""" + $ChangeListFoldersCount + """,""" + $SpecSheetFoldersCount + """,""" + $DrawingsCount + """")
    # Write-Host("$DIR,$EjectDrawingSheetFoldersCount,$ChangeListFoldersCount,$SpecSheetFoldersCount,$DrawingsCount")
    $activity = "図面TIFF内部　オートチェッカー"
    $status = "CSV書き込み中"
    $ProgressRate = [Math]::Round(($Counter / $DIRS.Length) * 100, 2, [MidpointRounding]::AwayFromZero)
    Write-Progress $activity $status -PercentComplete $ProgressRate -CurrentOperation "$ProgressRate % 完了"
    Start-Sleep -Milliseconds 50
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
