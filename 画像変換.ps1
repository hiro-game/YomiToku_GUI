#requires -Version 5.1
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
Add-Type -AssemblyName System.Windows.Forms

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    [System.Windows.MessageBox]::Show('このスクリプトは STA モードで実行してください。（powershell.exe -STA）','エラー')
    exit
}

#=== ImageMagick パス取得 ===#
function Get-MagickPath {
    $local = Join-Path $PSScriptRoot 'magick.exe'
    if (Test-Path $local) { return $local }
    return 'magick'
}

#=== ドロップされたパスからファイル一覧を取得 ===#
function Get-InputFiles {
    param([string[]]$Paths)

    $list = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($p in $Paths) {

        # ファイルとして存在するか LiteralPath で判定
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $list.Add([System.IO.FileInfo]$p)
            continue
        }

        # フォルダとして存在するか LiteralPath で判定
        if (Test-Path -LiteralPath $p -PathType Container) {

            foreach ($file in [System.IO.Directory]::EnumerateFiles(
                $p,
                "*",
                [System.IO.SearchOption]::AllDirectories
            )) {
                $list.Add([System.IO.FileInfo]$file)
            }
        }
    }

    return $list
}

#=== 数値正規化 ===#
function Normalize {
    param([string]$Text,[int]$Min,[int]$Max,[int]$Default)
    $v = 0
    if (-not [int]::TryParse($Text,[ref]$v)) { return $Default }
    if ($v -lt $Min) { return $Min }
    if ($v -gt $Max) { return $Max }
    return $v
}

#=== XAML ===#
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MagickSimpleConvert"
        Height="217" Width="520"
        AllowDrop="True" ResizeMode="CanMinimize">

    <Grid x:Name="MainGrid" Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- 1. 出力フォルダ -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <Label Content="出力先フォルダ:" ToolTip="変換後のファイルを保存するフォルダです。"/>
            <TextBox x:Name="OutputFolderBox" Width="320" Margin="5,0,5,0" IsReadOnly="True"
                     ToolTip="出力先フォルダを入力してください。"/>
            <Button x:Name="BrowseButton" Content="参照" Width="70"
                    ToolTip="出力先フォルダを選択します。"/>
        </StackPanel>

        <!-- 2. PDF/ベクター 読取解像度 & 拡張子フィルタ -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Label Content="PDF/ベクター読取解像度:" 
                   ToolTip="PDF やベクター画像の読み取り解像度(DPI)です。"/>
            <ComboBox x:Name="DensityBox" Width="80" Margin="5,0,15,0" IsEditable="True" Text="300"
                      ToolTip="PDF やベクター画像の読み取り解像度(DPI)です。">
                <ComboBoxItem Content="200"/>
                <ComboBoxItem Content="300"/>
                <ComboBoxItem Content="600"/>
                <ComboBoxItem Content="1200"/>
            </ComboBox>
            <Label Content="対象ファイル:" 
                   ToolTip="変換対象のファイル拡張子を指定します。空欄で全て対象。スペース区切りで複数指定可\n例： jpg png tif pdf"/>
            <TextBox x:Name="ExtFilterBox" Width="160" Margin="5,0,0,0"
                     ToolTip="スペース区切りで複数指定可 例： jpg png tif pdf\n空欄なら全てのファイルを変換します。"/>
        </StackPanel>

        <!-- 3. 出力形式ボタン -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
            <Label Content="出力形式:" ToolTip="出力する画像形式を JPG または PNG から選択します。"/>
            <Button x:Name="FormatButton" Width="80" Height="26" Margin="10,0,0,0"
                    Content="JPG" Background="Yellow"
                    ToolTip="クリックするたびに JPG / PNG が切り替わります。"/>
        </StackPanel>

        <!-- 4. JPG / PNG 設定 + 右下進捗表示 -->
        <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
        
            <!-- 左側：JPG / PNG 設定 -->
            <StackPanel Grid.Column="0" Orientation="Vertical">
                <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                    <Label Content="JPG 圧縮率 (1-100):"
                           ToolTip="数値が大きいほど高画質・大きなファイルになります。"/>
                    <TextBox x:Name="JpgQualityBox" Width="60" Margin="5,0,0,0" Text="90"
                             ToolTip="JPGの圧縮率を1〜100 の範囲で指定します。"/>
                </StackPanel>
        
                <StackPanel Orientation="Horizontal" Margin="0,0,0,0">
                    <Label Content="PNG 圧縮レベル:"
                           ToolTip="大きいほど圧縮率が高く、変換に時間が掛かります。"/>
                    <ComboBox x:Name="PngLevelBox"
                              Width="60"
                              Margin="5,0,15,0"
                              IsEditable="False">
                        <ComboBoxItem Content="0"/>
                        <ComboBoxItem Content="1"/>
                        <ComboBoxItem Content="2"/>
                        <ComboBoxItem Content="3"/>
                        <ComboBoxItem Content="4"/>
                        <ComboBoxItem Content="5"/>
                        <ComboBoxItem Content="6"/>
                        <ComboBoxItem Content="7"/>
                        <ComboBoxItem Content="8"/>
                        <ComboBoxItem Content="9"/>
                    </ComboBox>
        
                    <Label Content="PNG フィルター:"
                           ToolTip="5 は各行ごとに最適なフィルターを自動選択します。"/>
                    <ComboBox x:Name="PngFilterBox"
                              Width="60"
                              Margin="5,0,0,0"
                              IsEditable="False">
                        <ComboBoxItem Content="0"/>
                        <ComboBoxItem Content="1"/>
                        <ComboBoxItem Content="2"/>
                        <ComboBoxItem Content="3"/>
                        <ComboBoxItem Content="4"/>
                        <ComboBoxItem Content="5"/>
                    </ComboBox>
                </StackPanel>
            </StackPanel>
        
            <!-- 右側：進捗表示 -->
            <TextBlock Grid.Column="1"
                       x:Name="ProgressText"
                       Text="処理：0 / 入力：0"
                       VerticalAlignment="Bottom"
                       HorizontalAlignment="Right"
                       Margin="10,0,0,0"
                       FontSize="12"
                       Foreground="Gray"/>
        </Grid>
    </Grid>
</Window>
"@

#=== XAML 読み込み ===#
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$MainGrid        = $window.FindName('MainGrid')
$OutputFolderBox = $window.FindName('OutputFolderBox')
$BrowseButton    = $window.FindName('BrowseButton')
$DensityBox      = $window.FindName('DensityBox')
$ExtFilterBox    = $window.FindName('ExtFilterBox')
$FormatButton    = $window.FindName('FormatButton')
$JpgQualityBox   = $window.FindName('JpgQualityBox')
$PngLevelBox     = $window.FindName('PngLevelBox')
$PngFilterBox    = $window.FindName('PngFilterBox')
$ProgressText   = $window.FindName('ProgressText')

#=== 出力形式（内部状態） ===#
$script:TargetFormat = 'jpg'   # 'jpg' or 'png'

#=== 出力先フォルダ選択 ===#
$BrowseButton.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '出力先フォルダを選択してください'
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $OutputFolderBox.Text = $dlg.SelectedPath
    }
})

#=== 出力形式ボタン（JPG ⇄ PNG） ===#
$FormatButton.Add_Click({
    if ($script:TargetFormat -eq 'jpg') {
        $script:TargetFormat = 'png'
        $FormatButton.Content = 'PNG'
        $FormatButton.Background = 'LightBlue'
        $JpgQualityBox.IsEnabled = $false
        $PngLevelBox.IsEnabled  = $true
        $PngFilterBox.IsEnabled = $true
    } else {
        $script:TargetFormat = 'jpg'
        $FormatButton.Content = 'JPG'
        $FormatButton.Background = 'Yellow'
        $JpgQualityBox.IsEnabled = $true
        $PngLevelBox.IsEnabled  = $false
        $PngFilterBox.IsEnabled = $false
    }
})

# 初期状態：JPG
$script:TargetFormat = 'jpg'
$FormatButton.Content = 'JPG'
$FormatButton.Background = 'Yellow'
$JpgQualityBox.IsEnabled = $true
$PngLevelBox.IsEnabled  = $false
$PngFilterBox.IsEnabled = $false

# PNG 初期選択（7 / 5）
$PngLevelBox.SelectedIndex  = 7
$PngFilterBox.SelectedIndex = 5

#=== ドロップ処理 ===#
function Start-ProcessDrop {
    param($Paths)

    $outDir = $OutputFolderBox.Text
    if ([string]::IsNullOrWhiteSpace($outDir)) {
        [System.Windows.MessageBox]::Show('出力先フォルダが未設定です。','エラー')
        return
    }

    if (-not (Test-Path $outDir)) {
        try {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show('出力先フォルダを作成できませんでした。','エラー')
            return
        }
    }

    $files = Get-InputFiles -Paths $Paths
    if (-not $files -or $files.Count -eq 0) {
        [System.Windows.MessageBox]::Show('有効なファイルが見つかりませんでした。','情報')
        return
    }

    $total = $files.Count
    $done  = 0
    $ProgressText.Text = "処理：0 / 入力：$total"

    # density 正規化
    $densityText = $DensityBox.Text
    $density = Normalize $densityText 1 2400 300
    $DensityBox.Text = $density

    # 拡張子フィルタ（空欄なら全て対象）
    $extFilterRaw = $ExtFilterBox.Text
    $extFilterList = @()
    if (-not [string]::IsNullOrWhiteSpace($extFilterRaw)) {
        $extFilterList = $extFilterRaw.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) |
                         ForEach-Object { $_.Trim().ToLower().TrimStart('.') }
    }

    # JPG quality
    $jpgQ = Normalize $JpgQualityBox.Text 1 100 90
    $JpgQualityBox.Text = $jpgQ

    # PNG level/filter（ListBox 選択値）
    $pngLevel  = [int]$PngLevelBox.SelectedItem.Content
    $pngFilter = [int]$PngFilterBox.SelectedItem.Content

    $magick = Get-MagickPath
    $vectorExt = '.pdf','.ai','.eps','.svg','.ps'

    $success = 0
    $failed  = 0

foreach ($f in $files) {

    $src  = $f.FullName
    $ext  = [IO.Path]::GetExtension($src).ToLower()
    $extPlain = $ext.TrimStart('.')

    # 拡張子フィルタ適用（指定がある場合のみ）
    if ($extFilterList.Count -gt 0 -and -not ($extFilterList -contains $extPlain)) {
        continue
    }

    $base = [IO.Path]::GetFileNameWithoutExtension($src)
    $dest = Join-Path $outDir ($base + '.' + $script:TargetFormat)

    $args = @()

    # ベクター / PDF の場合は density を入力前に付与
    if ($vectorExt -contains $ext -or $ext -eq ".pdf") {
        $args += "-density"
        $args += $density
    }

# 入力側オプション（SVG / PDF のレンダリング精度）
$vectorExt = @(".svg", ".svgz", ".msvg", ".mvg", ".ai", ".eps", ".epsf", ".epsi", ".ps", ".ps2", ".ps3", ".wmf", ".emf")

if ($vectorExt -contains $ext -or $ext -eq ".pdf") {
    $args += "-density"
    $args += $density
}

# 入力ファイル（SVG の場合は msvg: を使う）
if ($ext -eq ".svg" -or $ext -eq ".svgz") {
    $args += "msvg:`"$src`""
} else {
    $args += "`"$src`""
}

# PNG の場合だけ出力側オプションを付ける
if ($script:TargetFormat -eq 'png') {

    # --- 出力側オプション ---
    if ($vectorExt -contains $ext) {
        # ベクター → PNG は透明化
        $args += "-background"
        $args += "none"
        $args += "-alpha"
        $args += "on"
    }
    elseif ($ext -eq ".pdf") {
        # PDF → PNG は透明を強制しない
        $args += "-alpha"
        $args += "on"
    }
    else {
        # ラスター → PNG は元の透明に従う
        $args += "-alpha"
        $args += "on"
    }

    # PNG 圧縮設定
    $pngQuality = "{0}{1}" -f $pngLevel,$pngFilter
    $args += "-quality"
    $args += $pngQuality

} else {
    # JPG の場合
    $args += "-quality"
    $args += $jpgQ
}

# 出力ファイル
$args += "`"$dest`""

# 実行
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $magick
$psi.Arguments = ($args -join ' ')
$psi.UseShellExecute = $false
$psi.RedirectStandardError = $true
$psi.CreateNoWindow = $true

Write-Host "MAGICK COMMAND:"
Write-Host "$magick $($psi.Arguments)"

$p = [System.Diagnostics.Process]::Start($psi)
$p.WaitForExit()

if ($p.ExitCode -eq 0) {
    $success++
} else {
    $failed++
}
$done++

$window.Dispatcher.Invoke([Action]{
    $ProgressText.Text = "処理：$done / 入力：$total"
}, 'Background')

}

    [System.Windows.MessageBox]::Show("変換完了`n成功: $success 件`n失敗: $failed 件",'完了')
}

#=== ウィンドウ全体をドロップ領域に ===#
$window.Add_Drop({
    param($sender,$e)

    # FileDrop は [] を含むファイル名を破壊するため使用しない
    if ($e.Data.GetDataPresent([Windows.DataFormats]::Text)) {

        $text = $e.Data.GetData([Windows.DataFormats]::Text)

        # 複数ファイルは改行区切りで来る
        $paths = $text -split "`r?`n" | ForEach-Object {
            $_.Trim()
        }

        Start-ProcessDrop -Paths $paths
        return
    }

    # 念のため FileDrop も残すが、Text を優先
    if ($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {

        $raw = $e.Data.GetData([Windows.DataFormats]::FileDrop)

        $paths = foreach ($p in $raw) {
            if ($p -is [System.Uri]) {
                $p.LocalPath
            } else {
                [string]$p
            }
        }

        Start-ProcessDrop -Paths $paths
    }
})

#=== 表示 ===#
$window.ShowDialog() | Out-Null
