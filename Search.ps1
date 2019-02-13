# �����t�H���_���w��ł���݌v�ɂȂ��Ă��܂�(�z��)�B

# System.Windows.Forms��L����
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# �K�v�ȏ���ݒ�
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = "�Ώۃf�B���N�g����I�����Ă��������B" 
$fbd.SelectedPath = "\\192.168.0.170\supersub\�}��Tiff�f�[�^"

# �_�C�A���O��\������
# $target = $fbd.ShowDialog() | Out-Null

# �I�����L�����Z�������ꍇ��NULL��Ԃ�
if ( $target -eq [System.Windows.Forms.DialogResult]::Cancel) {
    $targetPath = $null
}
else {
    $targetPath = $fbd.SelectedPath 
}

Write-Host $targetPath

# $FoldersConfigPath = $targetPath
$FoldersConfigPath = "\\192.168.0.170\supersub\�}��Tiff�f�[�^\TOYOTA"
$DIRS = (Get-ChildItem $FoldersConfigPath -Directory) -as [string[]]
# ���C���̏���
$errorMessage = ""
$resultMessage = ""
# Write-Host $DIRS.Contains("00007").ToString()
# Write-Host $DIRS.Contains("00000").ToString()

# ���s���̃p�X�擾/�ړ�
$path = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $path
$fileName = $path + "\Log.csv"
$file = New-Object System.IO.StreamWriter($fileName, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$file.WriteLine("""�H��"",""�o�A�t�H���_"",""�ύX���X�g�t�H���_"",""�d�l���t�H���_"",""�}�ʐ�""")
$Counter = 1
foreach ($DIR in $DIRS) {
    # $finderPath = ("FileSystem::$DIR") # dir $DIR/hoge.xlsx �ɂ���΁Ahoge.xlsx�t�@�C�������ɍi��܂�
    $finderPath = $FoldersConfigPath + "\" + $DIR
    $Folders = (Get-ChildItem $finderPath -Directory -Depth 0) -as [string[]]
    $Files = (Get-ChildItem $finderPath -File -Depth 0 -Name) -as [string[]]
    # �o�}�A����
    $EjectDrawingSheetFolders = ($Folders -match "�o�}�A����.*") -as [string[]]
    $EjectDrawingSheetFoldersCount = $EjectDrawingSheetFolders.Length
    # �ύX���X�g
    $ChangeListFolders = ($Folders -match "�ύX���X�g.*") -as [string[]]
    $ChangeListFoldersCount = $ChangeListFolders.Length
    # �d�l��
    $SpecSheetFolders = ($Folders -match "�d�l��.*") -as [string[]]
    $SpecSheetFoldersCount = $SpecSheetFolders.Length
    # �}��
    $Drawings = ($Files -match ".+\.tif?") -as [string[]]
    $DrawingsCount = $Drawings.Length
    

    $file.WriteLine("""" + $DIR + """,""" + $EjectDrawingSheetFoldersCount + """,""" + $ChangeListFoldersCount + """,""" + $SpecSheetFoldersCount + """,""" + $DrawingsCount + """")
    # Write-Host("$DIR,$EjectDrawingSheetFoldersCount,$ChangeListFoldersCount,$SpecSheetFoldersCount,$DrawingsCount")
    $activity = "�}��TIFF�����@�I�[�g�`�F�b�J�["
    $status = "CSV�������ݒ�"
    $ProgressRate = [Math]::Round(($Counter / $DIRS.Length) * 100, 2, [MidpointRounding]::AwayFromZero)
    Write-Progress $activity $status -PercentComplete $ProgressRate -CurrentOperation "$ProgressRate % ����"
    Start-Sleep -Milliseconds 50
    $Counter++
    if ($DIR -eq "00020") {
        # break
    }
}
$file.Close()


# �A�Z���u���̓ǂݍ���
Add-Type -Assembly System.Windows.Forms | Out-Null
#���ʕ\��
[void][System.Windows.Forms.MessageBox]::Show("Finish Research", "GF��掺")
Write-Host
Write-Host "Finished Please Input Any Key"
exit
