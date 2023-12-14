#################################################################################
# 処理名　｜IMEdictionarySearchTool
# 機能　　｜辞書ファイルの単語検索ツール
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜なし
#################################################################################
# 設定
# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み
#   フォーム用
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#   UI オートメーション
Add-Type -AssemblyName "UIAutomationClient"
Add-Type -AssemblyName "UIAutomationTypes"
$AutomationElement = [System.Windows.Automation.AutomationElement]
$TreeScope = [System.Windows.Automation.TreeScope]
$Condition = [System.Windows.Automation.Condition]
$InvokePattern = [System.Windows.Automation.InvokePattern]
$SendKeys = [System.Windows.Forms.SendKeys]
$Cursor = [System.Windows.Forms.Cursor]
# マウスの左クリック操作をおこなうための準備
$SendInputSource =@"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class MouseClick {
    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT {
        public int dx;
        public int dy;
        public int mouseData;
        public int dwFlags;
        public int time;
        public IntPtr dwExtraInfo;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public int type;
        public MOUSEINPUT mi;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    extern static uint SendInput(uint cInputs, INPUT[] pInputs, int cbSize);

    public static void Click() {
        INPUT[] input = new INPUT[2];
        input[0].mi.dwFlags = 0x0002;
        input[1].mi.dwFlags = 0x0004;
        SendInput(2, input, Marshal.SizeOf(input[0]));
    }
}
"@
Add-Type -TypeDefinition $SendInputSource -ReferencedAssemblies System.Windows.Forms, System.Drawing
$MouseClick = [MouseClick]
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = 'Stop'
# 定数
[System.String]$c_config_file = 'setup.ini'
# エラーコード enum設定
Add-Type -TypeDefinition @"
    public enum MESSAGECODE {
        Successful = 0,
        Abend,
        Cancel,
        Info_LoadedSettingfile,
        Confirm_ExecutionTool,
        Error_NotCore,
        Error_NotSupportedVersion,
        Error_NotWindows,
        Error_LoadingSettingfile,
        Error_NotExistsTargetpath,
        Error_EmptyTargetfolder,
        Error_EmptySearchkeywords,
        Error_NotMatchDatatype,
        Error_MaxRetries,
        Error_CountKeywordDictionary
    }
"@

### DEBUG ###
Set-Variable -Name "DEBUG_ON" -Value $false -Option Constant

### Function --- 開始 --->
#################################################################################
# 処理名　｜RemoveDoubleQuotes
# 機能　　｜先頭桁と最終桁にあるダブルクォーテーションを削除
#--------------------------------------------------------------------------------
# 戻り値　｜String（削除後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function RemoveDoubleQuotes {
    param (
        [System.String]$target_str
    )
    [System.String]$removed_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq '"') -and
            ($target_str.Substring($target_str.Length - 1, 1) -eq '"')) {
            # 先頭桁と最終桁のダブルクォーテーション削除
            $removed_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RemoveDoubleQuotes: target_str  [${target_str}]"
        Write-Host "                             removed_str [${removed_str}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $removed_str
}

#################################################################################
# 処理名　｜VerificationExecutionEnv
# 機能　　｜PowerShell環境チェック
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜なし
#################################################################################
Function VerificationExecutionEnv {
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_messages = ''

    # 環境情報を取得
    [System.Collections.Hashtable]$powershell_ver = $PSVersionTable

    # 環境の判定：Coreではない場合
    if ($powershell_ver.PSEdition -ne 'Core') {
        $messagecode = [MESSAGECODE]::Error_NotCore
        $messagecode_messages = RetrieveMessage $messagecode
        Write-Host $messagecode_messages -ForegroundColor DarkRed
    }
    # 環境の判定：メジャーバージョンが7より小さい場合
    elseif ($powershell_ver.PSVersion.Major -lt 7) {
        $messagecode = [MESSAGECODE]::Error_NotSupportedVersion
        $messagecode_messages = RetrieveMessage $messagecode
        Write-Host $messagecode_messages -ForegroundColor DarkRed
    }
    # 環境の判定：Windows OSではない場合
    elseif (-Not($IsWindows)) {
        $messagecode = [MESSAGECODE]::Error_NotWindows
        $messagecode_messages = RetrieveMessage $messagecode
        Write-Host $messagecode_messages -ForegroundColor DarkRed
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function VerificationExecutionEnv: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
#################################################################################
# 処理名　｜AcquisitionFormsize
# 機能　　｜Windowsフォーム用のサイズをモニターサイズから除算で設定
#--------------------------------------------------------------------------------
# 戻り値　｜String[]（変換後のサイズ：1要素目 横サイズ、2要素目 縦サイズ）
# 引数　　｜divisor: 除数（モニターサイズから除算するため）
#################################################################################
Function AcquisitionFormsize {
    param (
        [System.UInt32]$divisor
    )
    # 現在のモニターサイズを取得
    [Microsoft.Management.Infrastructure.CimInstance]$graphics_info = (Get-CimInstance -ClassName Win32_VideoController)
    [System.UInt32]$width = $graphics_info.CurrentHorizontalResolution
    [System.UInt32]$height = $graphics_info.CurrentVerticalResolution

    # モニターのサイズから除数で割る
    [System.UInt32]$form_width = $width / $divisor
    [System.UInt32]$form_height = $height / $divisor
    
    [System.UInt32[]]$form_size = @($form_width, $form_height)

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function AcquisitionFormsize: form_size [${form_size}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $form_size
}

#################################################################################
# 処理名　｜ConfirmYesno
# 機能　　｜YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno {
    param (
        [System.String]$prompt_message,
        [System.String]$prompt_title='実行前の確認'
    )

    # 除数「6」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(6)

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(($form_size[0] * 0.016), ($form_size[1] * 0.030))
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(($form_size[0] * 0.0156),($form_size[1] * 0.0285))
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.07))
    $label.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label.Text = $prompt_message
    $label.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = 'OK'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # フォーム表示
    [System.Boolean]$is_selected = ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: is_selected [${is_selected}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $is_selected
}

#################################################################################
# 処理名　｜ValidateInputValues
# 機能　　｜入力値の検証
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜setting_parameters[]
# 　　　　｜ - 項目01 作業フォルダー
# 　　　　｜ - 項目02 検索キーワード［複数の場合はカンマ（,）区切り］
# 　　　　｜ - 項目03 大文字と小文字を区別する
#################################################################################
Function ValidateInputValues {
    param (
        [System.String[]]$setting_parameters
    )
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful

    # メッセージボックス用
    [System.String]$messagebox_title = ''
    [System.String]$messagebox_messages = ''
    [System.String]$append_message = ''

    # 作業フォルダー
    #   入力チェック
    if ($setting_parameters[0] -eq '') {
        $messagecode = [MESSAGECODE]::Error_EmptyTargetfolder
        $messagebox_messages = RetrieveMessage $messagecode
        $messagebox_title = '入力チェック'
        ShowMessagebox $messagebox_messages $messagebox_title
    }
    #   存在チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        if (-Not(Test-Path $setting_parameters[0])) {
            $messagecode = [MESSAGECODE]::Error_NotExistsTargetpath
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
            "対象フォルダー: [$($setting_parameters[0])]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $messagebox_messages = RetrieveMessage $messagecode $append_message
            $messagebox_title = '存在チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    #   検索キーワード
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        # 入力チェック
        if ($setting_parameters[1] -eq '') {
            $messagecode = [MESSAGECODE]::Error_EmptySearchkeywords
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    #   大文字と小文字を区別する
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        # Boolean型のチェック
        #   Boolean型は値を代入する際に自動でデータ型のチェックが行われるので不要なチェックかも
        if (-Not([System.Boolean]::TryParse([System.String]$setting_parameters[2], [ref]$null))) {
            $messagecode = [MESSAGECODE]::Error_NotMatchDatatype
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: return [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　｜SettingInputValues
# 機能　　｜入力フォルダーの設定（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Object[]
# 　　　　｜ - 項目01 対象フォルダー        : 画面での設定値 - ツールの作業フォルダーとして使用
# 引数　　｜function_parameters[]
# 　　　　｜ - 項目01 ツール実行場所        : ツールの実行場所
# 　　　　｜ - 項目02 対象フォルダー        : 初期表示用の値 - 設定ファイルの設定値が反映
# 　　　　｜ - 項目03 検索キーワード        : 初期表示用の値 - 設定ファイルの設定値が反映
# 　　　　｜ - 項目04 大文字と小文字を区別  : 初期表示用の値 - 設定ファイルの設定値が反映
#################################################################################
Function SettingInputValues {
    param (
        [System.Object[]]$function_parameters
    )

    # 除数「3」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(3)

    # フォームの作成
    [System.String]$prompt_title = '実行前の設定'
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$($function_parameters[0])\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # 作業フォルダー - ラベル作成
    [System.Windows.Forms.Label]$label_input_folder = New-Object System.Windows.Forms.Label
    $label_input_folder.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.070))
    $label_input_folder.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_input_folder.Text = '作業場所として使用するフォルダーを指定してください。'
    $label_input_folder.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 作業フォルダー - テキストボックスの作成
    [System.Windows.Forms.TextBox]$textbox_input_folder = New-Object System.Windows.Forms.TextBox
    $textbox_input_folder.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.175))
    $textbox_input_folder.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75), 0)
    $textbox_input_folder.Text = $function_parameters[1]
    $textbox_input_folder.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 対象フォルダ― - 参照ボタンの作成
    [System.Windows.Forms.Button]$btnRefer = New-Object System.Windows.Forms.Button
    $btnRefer.Location = New-Object System.Drawing.Point(($form_size[0] * 0.820), ($form_size[1] * 0.175))
    $btnRefer.Size = New-Object System.Drawing.Size(75,25)
    $btnRefer.Text = '参照'

    # 検索キーワード - ラベル作成
    [System.Windows.Forms.Label]$label_search_keywords = New-Object System.Windows.Forms.Label
    $label_search_keywords.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.28))
    $label_search_keywords.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_search_keywords.Text = '検索キーワードの指定［ 複数の場合はカンマ（,）区切り ］'
    $label_search_keywords.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 検索キーワード - テキストボックスの作成
    [System.Windows.Forms.TextBox]$textbox_search_keywords = New-Object System.Windows.Forms.TextBox
    $textbox_search_keywords.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.385))
    $textbox_search_keywords.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75), 0)
    $textbox_search_keywords.Text = $function_parameters[2]
    $textbox_search_keywords.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 大文字と小文字を区別する - チェックボックスの作成
    [System.Windows.Forms.CheckBox]$checkbox_case_sensitive = New-Object System.Windows.Forms.CheckBox
    $checkbox_case_sensitive.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.49))
    $checkbox_case_sensitive.Size = New-Object System.Drawing.Size(20, 20)
    $checkbox_case_sensitive.Checked = $function_parameters[3]
    $checkbox_case_sensitive.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)
    
    # 大文字と小文字を区別する - ラベル作成
    [System.Windows.Forms.Label]$label_case_sensitive = New-Object System.Windows.Forms.Label
    $label_case_sensitive.Location = New-Object System.Drawing.Point(($form_size[0] *0.08),($form_size[1] * 0.49))
    $label_case_sensitive.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_case_sensitive.Text = '大文字と小文字を区別する'
    $label_case_sensitive.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = '次へ'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($label_input_folder)
    $form.Controls.Add($textbox_input_folder)
    $form.Controls.Add($btnRefer)
    $form.Controls.Add($label_search_keywords)
    $form.Controls.Add($textbox_search_keywords)
    $form.Controls.Add($checkbox_case_sensitive)
    $form.Controls.Add($label_case_sensitive)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # 参照ボタンの処理
    $btnRefer.add_click{
        #ダイアログを表示しファイルを選択する
        $folder_dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        if($folder_dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            $textbox_input_folder.Text = $folder_dialog.SelectedPath
        }
    }

    # フォーム表示
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.Int32]$max_retries = 3
    for ([System.Int32]$i=0; $i -le $max_retries; $i++) {
        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # 入力値のチェック
            [System.String[]]$setting_parameters = @()
            $setting_parameters = @(
                $textbox_input_folder.Text,
                $textbox_search_keywords.Text,
                $checkbox_case_sensitive.Checked
            )
            $messagecode = ValidateInputValues $setting_parameters

            # チェック結果が正常の場合
            if ($messagecode -eq [MESSAGECODE]::Successful) {
                $form = $null
                break
            }
        }
        else {
            $setting_parameters = @()
            $form = $null
            break
        }
        # 再試行回数を超過前の処理
        if ($i -eq $max_retries) {
            $messagecode = [MESSAGECODE]::Error_MaxRetries
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '再試行回数の超過'
            ShowMessagebox $messagebox_messages $messagebox_title
            $setting_parameters = @()
            $form = $null
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function SettingInputValues: setting_parameters [${setting_parameters}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $setting_parameters
}

#################################################################################
# 処理名　｜RetrieveMessage
# 機能　　｜メッセージ内容を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String（メッセージ内容）
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function RetrieveMessage {
    param (
        [MESSAGECODE]$target_code,
        [System.String]$append_message=''
    )
    [System.String]$return_messages = ''
    [System.String]$message = ''

    switch($target_code) {
        Successful                          {$message='正常終了';break}
        Abend                               {$message='異常終了';break}
        Cancel                              {$message='キャンセルしました。';break}
        Info_LoadedSettingfile              {$message='設定ファイルの読み込みが完了。';break}
        Confirm_ExecutionTool               {$message='ツールを実行します。';break}
        Error_NotCore                       {$message='PowerShellエディションが「 Core 」ではありません。';break}
        Error_NotSupportedVersion           {$message='PowerShellバージョンがサポート対象外です。（バージョン7未満）';break}
        Error_NotWindows                    {$message='実行環境がWindows OSではありません。';break}
        Error_LoadingSettingfile            {$message='設定ファイルの読み込み処理でエラーが発生しました。';break}
        Error_NotExistsTargetpath           {$message='所定の場所に設定ファイルがありません。';break}
        Error_EmptyTargetfolder             {$message='作業フォルダーが空で指定されています。';break}
        Error_EmptySearchkeywords           {$message='検索キーワードが空で指定されています。';break}
        Error_NotMatchDatatype              {$message='データ型と値があっていません。';break}
        Error_MaxRetries                    {$message='再試行回数を超過しました。';break}
        Error_CountKeywordDictionary        {$message='辞書ファイル内のキーワードをカウント実行時にエラーが発生しました。';break}
        default                             {break}
    }

    $sbtemp=New-Object System.Text.StringBuilder
    @("${message}`r`n",`
      "${append_message}`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $return_messages = $sbtemp.ToString()

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RetrieveMessage: return_messages [${return_messages}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $return_messages
}

#################################################################################
# 処理名　｜ShowMessagebox
# 機能　　｜メッセージボックスの表示
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function ShowMessagebox {
    param (
        [System.String]$messages,
        [System.String]$title,
        [System.String]$level='Information'
        # 指定可能なレベル一覧（$level）
        #   None
        #   Hand
        #   Error
        #   Stop
        #   Question
        #   Exclamation
        #   Waring
        #   Asterisk
        #   Information
    )

    [System.Windows.Forms.DialogResult]$dialog_result = [System.Windows.Forms.MessageBox]::Show($messages, $title, "OK", $level)
    
    switch($dialog_result) {
        {$_ -eq [System.Windows.Forms.DialogResult]::OK} {
            break
        }
    }
}
#################################################################################
# 処理名　｜CreateCountlists
# 機能　　｜キーワードを検索し件数を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String[]（検索キーワードとカウント数）
# 　　　　｜ - n次元目 項目01 検索キーワード
# 　　　　｜ - n次元目 項目02 キーワードの件数
# 引数　　｜targetfile   : 検索対象ファイル
# 　　　　｜keyword_lists: 検索キーワードリスト
# 　　　　｜casesensitive: 大文字・小文字を区別（true: 区別する、false: 区別しない）
#################################################################################
Function CreateCountlists {
    param (
        [System.String]$targetfile,
        [System.String[]]$keyword_lists,
        [System.Boolean]$casesensitive
    )
    [System.String]$keyword = ''
    [System.Int32]$keyword_count = 0
    [System.Object[]]$count_lists = @()

    # テキストファイル内の文字列チェック
    [System.String]$textdata = (Get-Content $targetfile)
    if ([string]::IsNullOrEmpty($textdata)){
        ### DEBUG ###
        if ($DEBUG_ON) {
            Write-Host '### DEBUG PRINT ###'
            Write-Host ''

            Write-Host "Function CreateCountlists: count_lists [${count_lists}]"

            Write-Host ''
            Write-Host '###################'
            Write-Host ''
            Write-Host ''
        }

        # 空のため早期リターン
        exit
    }

    # 複数キーワードで検索
    foreach($keyword in $keyword_lists) {
        if ($casesensitive) {
            # 大文字・小文字区別する
            $keyword_count = @(Select-String "${targetfile}" -Pattern "${keyword}" -CaseSensitive).Count
        }
        else {
            # 大文字・小文字区別しない
            $keyword_count = @(Select-String "${targetfile}" -Pattern "${keyword}").Count
        }

        $count_lists += ,@($keyword, $keyword_count)
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function CreateCountlists: count_lists [${count_lists}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $count_lists
}
#################################################################################
# 処理名　｜ShowCountlists
# 機能　　｜キーワード検索した結果を表示
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜count_lists: 検索キーワード毎に構成されたカウント数のリスト
# 　　　　｜targetfile : 検索対象ファイルのフルパス
#################################################################################
Function ShowCountlists {
    param (
        [System.Object[]]$count_lists,
        [System.String]$targetfile
    )
    # リスト内の文字列チェック
    if ([string]::IsNullOrEmpty($count_lists)){
        ### DEBUG ###
        if ($DEBUG_ON) {
            Write-Host '### DEBUG PRINT ###'
            Write-Host ''

            Write-Host "Function ShowCountlists: count_lists [${count_lists}]"

            Write-Host ''
            Write-Host '###################'
            Write-Host ''
            Write-Host ''
        }

        # 空のため早期リターン
        exit
    }

    # 配列内で最大のバイト数（Shift-JIS）を取得
    [System.Object[]]$to_bytes = [Management.Automation.PSSerializer]::DeSerialize([Management.Automation.PSSerializer]::Serialize($count_lists))
    [System.Int32]$i = 0
    [System.Int32]$max_length = 0
    for ($i = 0; $i -lt $to_bytes.Count; $i++) {
        $to_bytes[$i][0] = [System.Text.Encoding]::GetEncoding("shift_jis").GetByteCount($to_bytes[$i][0])
        if ($max_length -lt $to_bytes[$i][0]) {
            $max_length = $to_bytes[$i][0]
        }
    }

    # 複数キーワードで検索
    Write-Host ' ============ 検索キーワード と 件数 ============ '
    Write-Host ''
    Write-Host " 対象ファイル   [${targetfile}]"
    Write-Host ''
    Write-Host ' ------------------------------------------------ '
    Write-Host ''
    [System.Int32]$tab_count = 0
    [System.Int32]$tab_width = 4
    for ($i = 0; $i -lt $to_bytes.Count; $i++) {
        # 挿入するタブ数を計算
        $tab_count = [Math]::Ceiling(($max_length - [System.Int32]$to_bytes[$i][0]) / $tab_width)
        if ($tab_count -eq 0) {
            $tab_count = 1
        }

        if ($count_lists[$i][1] -eq 0) {
            Write-Host " 検索キーワード [$($count_lists[$i][0])]$("`t" * $tab_count)、件数 [$($count_lists[$i][1])件] "
        }
        else {
            Write-Host " 検索キーワード [$($count_lists[$i][0])]$("`t" * $tab_count)、件数 [$($count_lists[$i][1])件] " -ForegroundColor DarkRed
        }
    }
    Write-Host ''
    Write-Host ' ================================================ '
    Write-Host ''
    Write-Host ''
    Write-Host ''
}

#################################################################################
# 処理名　｜CountKeywordDictionary
# 機能　　｜辞書ファイル内のキーワードをカウント
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜targetfile      : カウント対象のファイル（テキスト形式の辞書ファイル）
# 　　　　｜serach_keywords   : カウントするキーワードのリスト
# 　　　　｜casesensitive_str   : 検索する際に大文字・小文字を区別するか（True：区別する、False：区別しない）
#################################################################################
Function CountKeywordDictionary {
    param (
        [System.String]$targetfile,
        [System.String]$serach_keywords,
        [System.Boolean]$casesensitive

    )
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_messages = ''

    [System.String[]]$keyword_lists = $serach_keywords.Split(',')
    [System.Object[]]$count_lists = @()

    # キーワードカウント
    try {
        $count_lists = CreateCountlists "${targetfile}" $keyword_lists $casesensitive

        ShowCountlists $count_lists $targetfile
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_CountKeywordDictionary
        $messagecode_messages = RetrieveMessage $messagecode
        Write-Host $messagecode_messages -ForegroundColor DarkRed
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function CountKeywordDictionary: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　｜CreateExportFolder
# 機能　　｜一時ファイルを格納するフォルダーを新規作成
#--------------------------------------------------------------------------------
# 戻り値　｜String（作成したフォルダー名。試行回数の超過もしくはエラーで作成できなかった場合は空文字を返す）
# 引数　　｜current_dir: 作業フォルダ―のパス
# 　　　　｜foldername : 対象フォルダー名
# 　　　　｜max_retries: 最大のリトライ回数
#################################################################################
Function CreateExportFolder {
    param (
        [System.String]$current_dir,
        [System.String]$foldername,
        [System.Int32]$max_retries=30
    )
    [System.String]$newfoldername = $foldername
    [System.Int32]$i = 0
    [System.String]$nowdate = (Get-Date).ToString("yyyyMMdd")
    [System.String]$number = ''
    for ($i=1; $i -le $max_retries; $i++) {
        # カウント数の数値を3桁で0埋めした文字列にする
        $number = "{0:000}" -f $i
        # 作成したいフォルダー名を生成
        $newfoldername = "${foldername}_${nowdate}-${number}"
        # 作成したいフォルダー名の存在チェック
        if (-Not (Test-Path "${current_dir}\${newfoldername}")) {
            break
        }

        # リトライ回数を超過し作成するフォルダー名を決定できなかった場合
        if ($i -eq $max_retries) {
            $newfoldername = ''
        }
    }

    [System.String]$newfolder_path = ''
    if ($newfoldername -ne '') {
        $newfolder_path = "${current_dir}\${newfoldername}"
        try {
            New-Item -Path "${newfolder_path}" -Type Directory > $null
        }
        catch {
            $newfolder_path = ''
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function CreateTempFolder: newfolder_path [${newfolder_path}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $newfolder_path
}

#################################################################################
# 処理名　｜GetElements
# 機能　　｜要素を取得する関数
# 　　　　｜* 参考URL <https://sqripts.com/2023/05/11/47993/>
#--------------------------------------------------------------------------------
# 戻り値　｜AutomationElement
# 引数　　｜なし
#################################################################################
function GetElements {
    Param($RootWindowName = $null)
    if ($null -eq $RootWindowName) {
        try {
            return $AutomationElement::RootElement.FindAll($TreeScope::Subtree, $Condition::TrueCondition)
        }
        catch {
            return $null
        }
    }
    else {
        $childrenElements = $AutomationElement::RootElement.FindAll($TreeScope::Children, $Condition::TrueCondition)
        foreach ($element in $childrenElements) {
            if ($element.GetCurrentPropertyValue($AutomationElement::NameProperty) -eq $RootWindowName) {
                return $element.FindAll($TreeScope::Subtree, $Condition::TrueCondition)
            }
        }
            Write-Host "指定された名前 '${RootWindowName}' のウィンドウが見つかりません。"
    }
    return $null
}
#################################################################################
# 処理名　｜FindElement
# 機能　　｜要素を検索する関数
# 　　　　｜* 参考URL <https://sqripts.com/2023/05/11/47993/>
#--------------------------------------------------------------------------------
# 戻り値　｜element
# 引数　　｜なし
#################################################################################
function FindElement {
    Param($RootWindowName = $null, $PropertyType, $Identifier, $Timeout)
    $startTime = (Get-Date).Ticks
    do {
        foreach ($element in GetElements -RootWindowName $RootWindowName) {
            try {
                if ($element.GetCurrentPropertyValue($AutomationElement::$PropertyType) -eq $Identifier) {
                    return $element
                }
            }
            catch {
                continue
            }
        }
    }
    while (((Get-Date).Ticks - $startTime) -le ($Timeout * 10000))
    throw "指定された要素 '${Identifier}' が見つかりません。"
}
#################################################################################
# 処理名　｜ClickElement
# 機能　　｜クリック操作をおこなう関数
# 　　　　｜* 参考URL <https://sqripts.com/2023/05/11/47993/>
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜なし
#################################################################################
function ClickElement {
    Param($RootWindowName = $null, $PropertyType, $Identifier, $Timeout = 5000)
    $startTime = (Get-Date).Ticks
    do {
        $element = FindElement -RootWindowName $RootWindowName -PropertyType $PropertyType -Identifier $Identifier -Timeout $Timeout
        $isEnabled = $element.GetCurrentPropertyValue($AutomationElement::IsEnabledProperty)
        if ($isEnabled -eq "True") { break }
    }
    while (((Get-Date).Ticks - $startTime) -le ($Timeout * 10000))
    if ($isEnabled -ne "True") {
        throw "指定された要素 '${Identifier}' が有効状態になりません。"
    }

    if ($element.GetCurrentPropertyValue($AutomationElement::IsInvokePatternAvailableProperty) -eq "True") {
        $element.GetCurrentPattern($InvokePattern::Pattern).Invoke()
    }
    else {
        # IsInvokePatternAvailablePropertyがFalseの時はマウスカーソルを要素に移動して左クリックする
        $clickablePoint = $element.GetClickablePoint()
        $Cursor::Position = New-Object System.Drawing.Point($clickablePoint.X, $clickablePoint.Y)
        $MouseClick::Click()
    }
}
#################################################################################
# 処理名　｜SendKeys
# 機能　　｜キーボード操作をおこなう関数
# 　　　　｜* 参考URL <https://sqripts.com/2023/05/11/47993/>
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜なし
#################################################################################
function SendKeys {
    Param($RootWindowName = $null, $PropertyType, $Idendifier = $null, $Keys, $Timeout = 5000)
    if ($null -ne $Idendifier) {
        $element = FindElement -RootWindowName $RootWindowName -PropertyType $PropertyType -Identifier $Idendifier -Timeout $Timeout
        $element.SetFocus()
    }
    $SendKeys::SendWait($Keys)
}
#################################################################################
# 処理名　｜SaveWebCaptureByMicrosoftEdge
# 機能　　｜Microsoft EdgeでWebキャプチャの保存操作をおこなう関数
# 　　　　｜* 参考URL <https://sqripts.com/2023/05/11/47993/>
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜なし
#################################################################################
function SaveWebCaptureByMicrosoftEdge {
    SendKeys -Keys "^(+S)"
    ClickElement -PropertyType "AutomationIdProperty" -Identifier "view_52561"
    ClickElement -PropertyType "AutomationIdProperty" -Identifier "save_button_id"
    Start-Sleep -Seconds 3
    SendKeys -Keys "{ESCAPE}"
}
#################################################################################
# 処理名　｜ExportDictionaryfile
# 機能　　｜テキスト形式で辞書ファイルを出力
# 　　　　｜
#--------------------------------------------------------------------------------
# 戻り値　｜String（出力した辞書ファイルのフルパス）
# 引数　　｜setting_parameters[]
# 　　　　｜ - 項目01 “単語の登録”ウィンドウの実行ファイルのパス
# 　　　　｜ - 項目02 作業フォルダーのパス
# 　　　　｜ - 項目03 エクスポート用のフォルダー名（作業フォルダ―配下に作成）
# 　　　　｜ - 項目04 エクスポート用のファイル名（作業フォルダ―のエクスポート用フォルダー配下にエクスポート）
#################################################################################
function ExportDictionaryfile {
    param (
        [System.String[]]$setting_parameters
    )

    [System.String]$export_textfile = ''

    try {
        # 単語の登録ウィンドウの起動
        Start-Process $setting_parameters[0]

        # テキスト形式の辞書ファイルを格納するフォルダーの作成
        [System.String]$output_folder_path = CreateExportFolder $setting_parameters[1] $setting_parameters[2]
        $export_textfile = "$output_folder_path\$($setting_parameters[3])"

        # 単語の登録ウィンドウからユーザー辞書ツールを起動
        Start-Sleep -Milliseconds $waitMilliseconds
        ClickElement -RootWindowName '単語の登録' -PropertyType 'AutomationIdProperty' -Identifier '658'

        # ユーザー辞書ツールで Alt + T を入力し メニューバー「ツール(T)」 を選択
        Start-Sleep -Milliseconds $waitMilliseconds
        SendKeys -RootWindowName 'Microsoft IME ユーザー辞書ツール' -PropertyType 'AutomationIdProperty' -Keys '%t'

        # ユーザー辞書ツールのメニューバー“ツール(T)”で  P を入力し 「一覧の出力(P)」 を選択
        Start-Sleep -Milliseconds $waitMilliseconds
        SendKeys -RootWindowName 'Microsoft IME ユーザー辞書ツール' -PropertyType 'AutomationIdProperty' -Keys 'p'

        Start-Sleep -Milliseconds $waitMilliseconds
        SendKeys -RootWindowName 'ファイル名(N):' -PropertyType 'AutomationIdProperty' -Identifier '1148' -Keys $export_textfile

        Start-Sleep -Milliseconds $waitMilliseconds
        # ClickElement -RootWindowName '一覧の出力:単語一覧' -PropertyType 'AutomationIdProperty' -Identifier '1'
        # 上記、引数のRootWindowNameでウィンドウを指定する方法だと動作しなかったので下記に変更
        SendKeys -RootWindowName '一覧の出力:単語一覧' -PropertyType 'AutomationIdProperty' -Keys '%s'

        Start-Sleep -Milliseconds $waitMilliseconds
        # ClickElement -RootWindowName '一覧の出力' -PropertyType 'AutomationIdProperty' -Identifier '673'
        # 上記、引数のRootWindowNameでウィンドウを指定する方法だと動作しなかったので下記に変更
        SendKeys -RootWindowName '一覧の出力' -PropertyType 'AutomationIdProperty' -Keys '{ENTER}'

        Start-Sleep -Milliseconds $waitMilliseconds
        SendKeys -RootWindowName 'Microsoft IME ユーザー辞書ツール' -PropertyType 'AutomationIdProperty' -Keys '%{F4}'
    }
    catch {
        $export_textfile = ''
    }
    
    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ExportDictionaryfile: export_textfile [${export_textfile}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }
    
    return $export_textfile
}
### Function <--- 終了 ---

### Main process --- 開始 --->
#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
# 初期設定
#   メッセージ関連
[MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
[System.String]$prompt_title = ''
[System.String]$prompt_message = ''
[System.String]$messagecode_messages = ''
[System.String]$append_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# ボタンをクリックした後の待機ミリ秒を設定する
$waitMilliseconds = 300

# PowerShell環境チェック
$messagecode = VerificationExecutionEnv

# 設定ファイルの読み込み
if ($messagecode -eq [MESSAGECODE]::Successful) {
    # ディレクトリの取得
    [System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
    Set-Location $current_dir'\..\..'
    [System.String]$root_dir = (Convert-Path .)

    # Configファイルのフルパスを作成  
    $sbtemp=New-Object System.Text.StringBuilder
    @("${current_dir}",`
      '\',`
      "${c_config_file}")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    [System.String]$config_fullpath = $sbtemp.ToString()

    # 読み込み処理
    try {
        [System.Collections.Hashtable]$config = (Get-Content $config_fullpath -Raw -Encoding UTF8).Replace('\','\\') | ConvertFrom-StringData

        # 変数に格納
        [System.String]$CONFIG_DICTSEARCH_WORKFOLDER_PATH=RemoveDoubleQuotes($config.dictsearch_workfolder_path)
        [System.String]$CONFIG_DICT_SEARCH_KEYWORDS=(RemoveDoubleQuotes($config.dict_search_keywords))
        [System.Boolean]$CONFIG_DICT_CASE_SENSITIVE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.dict_case_sensitive)))
        [System.Boolean]$CONFIG_AUTO_MODE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.auto_mode)))
        [System.String]$CONFIG_EXPORT_FOLDERNAME=RemoveDoubleQuotes($config.export_foldername)
        [System.String]$CONFIG_TEMP_FILENAME=RemoveDoubleQuotes($config.temp_filename)

        # 通知
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
        "対象ファイル: [${config_fullpath}]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_LoadedSettingfile) $append_message
        Write-Host $prompt_message
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_LoadingSettingfile
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
          "エラーの詳細: [${config_fullpath}$($_.Exception.Message)]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $messagecode_messages = RetrieveMessage $messagecode $append_message
    }   
}

# 入力値の設定・検証
if ($messagecode -eq [MESSAGECODE]::Successful) {
    [System.Object[]]$function_parameters = @()
    [System.Object[]]$setting_parameters = @()
    if (-Not($CONFIG_AUTO_MODE)) {
        # 対話式の場合：入力値を画面で設定（入力値の検証を含む）
        [System.String]$input_folder = $CONFIG_DICTSEARCH_WORKFOLDER_PATH
        if ($input_folder -eq '') {
            $input_folder = $current_dir
        }
        $function_parameters = @(
            $root_dir,
            $input_folder,
            $CONFIG_DICT_SEARCH_KEYWORDS,
            $CONFIG_DICT_CASE_SENSITIVE
        )
        
        $setting_parameters = SettingInputValues $function_parameters
        if ($null -eq $setting_parameters) {
            $messagecode = [MESSAGECODE]::Cancel
        }
    }
    else {
        # 自動実行の場合：入力値の検証のみ
        #   入力値のチェック
        [System.String[]]$setting_parameters = @()
        $setting_parameters = @(
            $CONFIG_DICTSEARCH_WORKFOLDER_PATH,
            $CONFIG_DICT_SEARCH_KEYWORDS,
            $CONFIG_DICT_CASE_SENSITIVE
        )
        $messagecode = ValidateInputValues $setting_parameters
    }
}

# 実行有無の確認
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $prompt_message = RetrieveMessage ([MESSAGECODE]::Confirm_ExecutionTool)
    If (ConfirmYesno $prompt_message) {
        # テキスト形式で辞書ファイルを出力
        $function_parameters = @(
            "$Env:windir\system32\IME\IMEJP\IMJPDCT.EXE",   # 単語の登録ウィンドウ（IMJPDCT.EXE）の絶対パス
            $CONFIG_DICTSEARCH_WORKFOLDER_PATH,             # テキスト形式の辞書ファイルを出力する作業フォルダ―
            $CONFIG_EXPORT_FOLDERNAME,                      # 作業フォルダ―配下に作成するエクスポートするフォルダー名
            $CONFIG_TEMP_FILENAME                           # テキスト形式の辞書ファイルの名前
        )
        
        [System.String]$export_textfile = ExportDictionaryfile $function_parameters

        if ($export_textfile -eq '') {
            $messagecode = [MESSAGECODE]::Error_ExportDictionaryfile
        }
    }
}

# キーワード検索の実行＆カウント結果の表示
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $messagecode = CountKeywordDictionary $export_textfile $CONFIG_DICT_SEARCH_KEYWORDS $CONFIG_DICT_CASE_SENSITIVE
}

#   処理結果の表示
[System.String]$append_message = ''
$sbtemp=New-Object System.Text.StringBuilder
if ($messagecode -eq [MESSAGECODE]::Successful) {
    @("`r`n",`
      "メッセージコード: [${messagecode}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_messages = RetrieveMessage $messagecode $append_message
    Write-Host $messagecode_messages
}
else {
    @("`r`n",`
      "メッセージコード: [${messagecode}]`r`n",`
      $messagecode_messages)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_messages = RetrieveMessage ([MESSAGECODE]::Abend) $append_message
    Write-Host $messagecode_messages -ForegroundColor DarkRed
}

# 終了
exit $messagecode
### Main process <--- 終了 ---
