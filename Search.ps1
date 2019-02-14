# �����t�H���_���w��ł���݌v�ɂȂ��Ă��܂�(�z��)�B

# System.Windows.Forms��L����
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# ��{�ݒ�
$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
$fbd.Description = "�q�於�t�H���_��I�����Ă��������B" 
$fbd.SelectedPath = "\\192.168.0.170\supersub\�}��Tiff�f�[�^"
$TargetFolders = @("�o�}�A����", "�ύX���X�g", "�d�l��")

# �_�C�A���O��\������
$target = $fbd.ShowDialog() | Out-Null
# �I�����L�����Z�������ꍇ��NULL��Ԃ�
if ( $target -eq [System.Windows.Forms.DialogResult]::Cancel) {
    $targetPath = $null
}
else {
    $targetPath = $fbd.SelectedPath 
}
Write-Host $targetPath
$FoldersConfigPath = $targetPath
# $FoldersConfigPath = "\\192.168.0.170\supersub\�}��Tiff�f�[�^\TOYOTA"
$DIRS = (Get-ChildItem $FoldersConfigPath -Directory) -as [string[]]

# ���C���̏���
# ���s���̃p�X�擾/�ړ�
$path = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $path
$FoldersConfigPathSplit = $FoldersConfigPath.Split("\")
$fileName = $path + "\log\$($FoldersConfigPathSplit[$FoldersConfigPathSplit.Length-1]).csv"
$file = New-Object System.IO.StreamWriter($fileName, $false, [System.Text.Encoding]::GetEncoding("sjis"))
$file.Write("�H��,�}�ʐ�")
foreach ($item in $TargetFolders) {
    $file.Write(",$($item)�t�H���_,$($item)��")
}
$file.WriteLine("")
$Counter = 1
foreach ($DIR in $DIRS) {
    # $finderPath = ("FileSystem::$DIR") # dir $DIR/hoge.xlsx �ɂ���΁Ahoge.xlsx�t�@�C�������ɍi��܂�
    $finderPath = $FoldersConfigPath + "\" + $DIR
    $Folders = (Get-ChildItem $finderPath -Directory -Depth 0) -as [string[]]
    $Files = (Get-ChildItem $finderPath -File -Depth 0 -Name) -as [string[]]
    # �}��
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
    $activity = "�}��TIFF�����@�I�[�g�`�F�b�J�["
    $status = "CSV�������ݒ�"
    $ProgressRate = [Math]::Round(($Counter / $DIRS.Length) * 100, 2, [MidpointRounding]::AwayFromZero)
    Write-Progress $activity $status -PercentComplete $ProgressRate -CurrentOperation "$ProgressRate%����"
    Start-Sleep -Milliseconds 10
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
