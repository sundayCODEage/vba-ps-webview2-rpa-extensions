Attribute VB_Name = "Mod_TestRun3_Base拡張機能"
Option Explicit

' ==============================================================================
' 【高精度タイマー用 API宣言】
' .. Win32 APIを使用してミリ秒未満（マイクロ秒単位）の処理時間を正確に測定する
' ==============================================================================
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' --- モジュールレベル変数 ---
Private freq As Currency             ' CPUの動作周波数 (タイマー精度)
    ' タイマー用変数
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double

Private rpaEngine As Ps_Engine       ' RPA操作エンジン本体クラス

' --- API宣言 ---
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- エンジンパス等の定数定義 ---
Private Const PS_ENGINE_LIB As String = "\Ps_Engine_Core_v303.ps1"
    Const cTimeout20    As Integer = 20 ' Psモジュール：デフォルト(10)
    ' 公開サイト（特に、the-internet）が繋がりにくい時が有る Wait-WebPageLoadのタイムアウトを指定する。(通常は不要)

' ==============================================================================
'   RPAエンジン テスト3 (拡張機能)
' ==============================================================================
Sub Test_Run3_Base()

    Dim enginePath As String
    Dim sessionId As String
    Dim useCdpPort As Integer   ' 通信モードの設定 (0: 標準, 9222: CDP高速通信)

    ' CPUのタイマー周波数を取得 (高精度計測の準備)
    QueryPerformanceFrequency freq
    
    ' --- 実行環境・セッションの初期化 ---
    enginePath = ThisWorkbook.Path & PS_ENGINE_LIB
    sessionId = "SESSION_" & Format(Now, "yyyyMMdd_HHmmss")
    ' --- CDP通信モードの設定 ---
    ' 通常の運用では標準モード(0)を推奨しますが、本テストシナリオでは
    ' CDP機能(Invoke-CdpScript等)を検証するため、ポート番号(例: 9222)を指定します。
    Debug.Print " >> 通常の運用では標準モード(0)を推奨、テストではポート番号を指定します。"
    useCdpPort = 9222
    
    Set rpaEngine = New Ps_Engine
    
    ' --- PowerShellエンジンを起動 ---
    ' 第3引数: CDPポート番号 (9222)
    ' 第4引数: IsDebugModeFlg (True = 実行時のパラメータや詳細ログをSTDOUTへ出力する)
    If Not rpaEngine.StartEngine(sessionId, enginePath, useCdpPort, True) Then
        MsgBox "RPAエンジンの起動に失敗しました。", vbCritical, "起動エラー"
        Exit Sub
    End If

    ' ==========================================================================
    ' --- テストシナリオの呼び出し ---
    ' ==========================================================================
    ' Excelを強制的に最前面に持ってくる
    Call ForceFocusExcel

    Dim prompt As String
    Dim ans As String

    prompt = "パートを選んでください：" & vbCrLf & _
            " 0. ** 全パート( 99:除く )一括テスト **" & vbCrLf & _
            " ---" & vbCrLf & _
            " 22. 仮想タブ（ポップアップ）でのサイレントDL自動継承" & vbCrLf & _
            " 23. 業務システム複合UI拡張＆オーバーライト・帳票DL" & vbCrLf & _
            " ---" & vbCrLf & _
            " 99. 【 TEST 】"
  
    ans = InputBox(prompt, "テストパート選択")
    
    If ans = "" Then
        MsgBox "キャンセルされました。", vbInformation
        GoTo CleanUp
    Else
        MsgBox "選択したテストパート: ( " & ans & " )", vbInformation
    End If

    ' 裏に隠れてしまったブラウザ（RPA Browser）を一番手前に呼び戻す
    On Error Resume Next
    AppActivate "RPA Browser"
    On Error GoTo 0
    ' 画面の切り替えが完全に落ち着くまでn秒待つ
    Application.Wait Now + TimeValue("00:00:01")

    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    ' .. (有効関数: Wait-WebXPathElement、Invoke-WebXPathClick、Set-WebXPathTextInput、Get-WebXPathText)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)

    ' クラスモジュール： SendAndReceive-タイムアウトを確認する。
    Debug.Print " ◆ SendAndReceive-タイムアウト: " & rpaEngine.Loop_TimeoutSec
    rpaEngine.Loop_TimeoutSec = 59  ' クラスモジュール初期値： 45

    ' --- テスト実行 ---
    Select Case ans
        Case "0"
            If Not TEST_Part22() Then GoTo CleanUp
            Sleep 2000 ' テスト画面を目視する（以下、Sleepも同様、適時調整する）
            If Not TEST_Part23() Then GoTo CleanUp
            Sleep 2000
            If Not TEST_Part23() Then GoTo CleanUp
        Case "22"
            If TEST_Part22() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "23"
            If TEST_Part23() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case "99" ' テスト用
            If TEST_Part99() Then MsgBox "テストが終了しました。", vbInformation
            GoTo CleanUp
        Case Else
            MsgBox "無効な入力です。", vbExclamation
            GoTo CleanUp
    End Select

    On Error GoTo ErrorHandler
    ' ==========================================================================
    Debug.Print "=== テスト完了 ==="
    MsgBox "テストが完了しました！" & vbCrLf & _
           "Logsフォルダに出力された証跡を確認してください。", vbInformation

CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Application.StatusBar = False ' ステータスバーのリセット
    Exit Sub
    
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        Debug.Print " 詳細(Details): " & rpaErr.Details
        
        MsgBox "RPAエンジン実行中にエラーが発生しました。" & vbCrLf & _
               "関数: " & rpaErr.FunctionName & vbCrLf & _
               "内容: " & rpaErr.Message, vbCritical, "RPAシステムエラー"
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If
    
    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_ErrorDump")
    Resume CleanUp
End Sub

' ==========================================================================
' 共通エラーハンドラ (サブプロシージャ用)
' ==========================================================================
Private Sub HandleError_Type1(ByVal procedureName As String)
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)

        Debug.Print "【発生プロシージャ】 : " & procedureName
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        Debug.Print " 詳細(Details): " & rpaErr.Details

        MsgBox "RPAエンジン実行中にエラーが発生しました。" & vbCrLf & _
               "箇所: " & procedureName & vbCrLf & _
               "関数: " & rpaErr.FunctionName & vbCrLf & _
               "内容: " & rpaErr.Message, vbCritical, "RPAシステムエラー"
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If

    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_ErrorDump_" & procedureName)
End Sub

Private Function TEST_Part22() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 22: 仮想タブ（ポップアップ）でのサイレントDL自動継承テスト"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim sandboxPath As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String
    Dim targetHref As String
    Dim ext As String

    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    If Dir(saveDir, vbDirectory) = "" Then MkDir saveDir

    ' ==========================================================================
    ' 【TEST 1】 公開サイト (herokuapp) における target="_blank" 仮想タブDL検証
    ' ダウンロードリンクに動的に target="_blank" 属性を付与してからクリックし、
    ' 新規仮想タブ（または別ウィンドウ要求）経由でも Enable-SilentDownload の設定が
    ' 継承されて指定ファイル名でサイレント保存されるかを検証する。
    ' ==========================================================================
    Debug.Print "TEST1: >> 公開サイト (herokuapp/download) で target='_blank' 経由のDL検証..."
    targetUrl = "https://the-internet.herokuapp.com/download"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad", CreateParams("TimeoutSec", cTimeout20)

    ' 1番目のリンクから拡張子を取得し、保存ファイル名を決定
    targetHref = rpaEngine.RunAction("Get-WebAttribute", CreateParams("XPath", "//div[@class='example']/a[1]", "Attribute", "href"))
    If InStrRev(targetHref, ".") > 0 Then
        ext = Mid(targetHref, InStrRev(targetHref, "."))
    Else
        ext = ".bin"
    End If

    saveFileName = "PopupPublic_" & Format(Now, "HHmmss") & ext
    fullSavePath = saveDir & "\" & saveFileName
    If Dir(fullSavePath) <> "" Then Kill fullSavePath

    ' 1. 親画面(P01)でサイレントダウンロードを事前予約
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)

    ' 2. 1番目のダウンロードリンクに target="_blank" を付与（別ウィンドウ要求を強制発生させる）
    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", _
        "const a = document.querySelector('.example a'); if(a) { a.setAttribute('target', '_blank'); } return true;")

    QueryPerformanceCounter startTime
    ' 3. リンクをクリックし、ファイル保存完了を待機
    rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//div[@class='example']/a[1]")
    result = rpaEngine.RunAction("Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/公開サイト仮想タブDL (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    If Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ [OK] 保存/公開サイト仮想タブDL: " & fullSavePath
    Else
        Err.Raise 999, "Test", "公開サイトでの仮想タブDLファイルが見つかりません"
    End If

    ' 念のため親タブ(P01)をアクティブに戻しておく
    rpaEngine.RunAction "Switch-Tab", CreateParams("TabId", "P01")
    Sleep 1000

    ' ==========================================================================
    ' 【TEST 2】 ローカル sandbox (13_popup_download.html) による業務システム再現検証
    ' 業務システムで頻出する「ボタンを押すと window.open で子画面（仮想タブ）が開き、
    ' 子画面側で帳票（CSV等）の自動ダウンロードが発火した後、window.close で自動的に閉じて
    ' 親画面へ復帰する」一連のライフサイクルを検証する。
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST2: >> ローカル (14_popup_dl_and_custom_ui.html) で子画面自動DL＆自動クローズ検証..."
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")

    targetUrl = "file:///" & sandboxPath & "14_popup_dl_and_custom_ui.html"
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    saveFileName = "PopupLocal_Report_" & Format(Now, "HHmmss") & ".csv"
    fullSavePath = saveDir & "\" & saveFileName
    If Dir(fullSavePath) <> "" Then Kill fullSavePath

    ' 1. 親画面(P01)で保存ファイル名を更新予約
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)

    QueryPerformanceCounter startTime
    ' 2. 「帳票出力」ボタンをクリック（→ 仮想タブ生成 → 自動DL発火 → 自動window.close）
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#btnPopupDl")

    ' 3. 仮想タブ側で発火したダウンロードの完了を待機
    result = rpaEngine.RunAction("Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/ローカル仮想タブ自動DL (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    If Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ [OK] 保存/ローカル仮想タブDL: " & fullSavePath
    Else
        Err.Raise 999, "Test", "ローカル仮想タブDLファイルが見つかりません"
    End If

    ' 4. 子ウィンドウの自動クローズ(800ms後)と親タブ(P01)への自動復帰を確認
    Sleep 1200
    result = rpaEngine.RunAction("List-Tabs")
    Debug.Print " ◆ 自動クローズ後のタブ一覧: " & result
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#statusMsg"))
    Debug.Print " ◆ 親画面(P01)のステータステキスト: " & result

    TEST_Part22 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名)
    Call HandleError_Type1("TEST_Part22")
    TEST_Part22 = False
End Function

Private Function TEST_Part23() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 23: 業務システム複合UI拡張 ＆ オーバーライト・帳票DL統合テスト"
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler

    Dim result As String
    Dim targetUrl As String
    Dim sandboxPath As String
    Dim saveDir As String
    Dim saveFileName As String
    Dim fullSavePath As String

    saveDir = ThisWorkbook.Path & "\RPA_Downloads"
    If Dir(saveDir, vbDirectory) = "" Then MkDir saveDir

    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
    targetUrl = "file:///" & sandboxPath & "14_popup_dl_and_custom_ui.html"
    ' ページ遷移 ＆ DOM構築完了の待機
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebPageLoad"

    ' ==========================================================================
    ' 【TEST 1】 標準関数（オーバーライト前）での「壁（未反映・エラー）」の再現検証
    '  ※ 手動で Lib-WebCustomUI_v303.ps1 の「_Custom」を消してオーバーライトした際は、
    '     この TEST 1 自体もエラーにならず一発で成功へと変化します。
    ' ==========================================================================
    Debug.Print "TEST1: >> 標準関数 (Set-WebTextInput / Invoke-WebClick) でのエラー再現検証..."

    ' 1. 標準の Set-WebTextInput で開始日を入力（DateTagのonfocusを通らないため後でリセットされる）
    rpaEngine.RunAction "Set-WebTextInput", CreateParams( _
        "Selector", "#dsp_IP999SC\~tfdKaisiBi", _
        "Value", "5080401")

    ' 2. ピッカーが取得した #HJ-SET-FOCUSFUNCTION11 を標準の Invoke-WebClick でクリック
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#HJ-SET-FOCUSFUNCTION11")
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#f11_result"))
    Debug.Print " ◆ [比較前] #HJ-SET-FOCUSFUNCTION11 標準クリック結果: " & result

    ' 3. 親の #FUNC11 を標準クリックしても、開始日が初期値(5081101)にリセットされてしまうことを確認
    rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", "#FUNC11")
    result = rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#f11_result"))
    Debug.Print " ◆ [比較前] 標準入力後の #FUNC11 送信結果          : " & result
    Debug.Print " ◆ [比較前] 開始日の表示値（初期値へリセット）     : " & _
        rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#dsp_IP999SC\~tfdKaisiBi"))

    ' ==========================================================================
    ' 【TEST 2】 専用拡張関数 (Set-WebIpkDateInput / Set-WebPairedInput) の検証
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST2: >> 専用拡張関数による2層構造日付・readonlyペア要素の更新検証..."

    result = rpaEngine.RunAction("Set-WebIpkDateInput", CreateParams( _
        "Selector", "#dsp_IP999SC\~tfdKaisiBi", _
        "Value", "5080401"))
    Debug.Print " ◆ [拡張関数] 開始日更新JSON : " & result

    result = rpaEngine.RunAction("Set-WebPairedInput", CreateParams( _
        "Selector", "#dsp_IP999SC\~tfdSyuryoBi", _
        "Value", "5090331", _
        "DisplayValue", "令和09年03月31日"))
    Debug.Print " ◆ [拡張関数] 終了日更新JSON : " & result

    ' ==========================================================================
    ' 【TEST 3】 オーバーライト検証関数 (_Custom) ＆ 仮想タブ帳票CSVダウンロード融合検証
    '  ピッカーで取得したままのセレクタ (#dsp_..., #HJ-SET-FOCUSFUNCTION11) を渡し、
    '  内部ルーティングによる日付確定 → F11発火 → 仮想タブCSV保存 → 自動クローズを検証する。
    ' ==========================================================================
    Debug.Print "=========================================================="
    Debug.Print "TEST3: >> オーバーライト関数 (_Custom) ＋ 仮想タブCSVサイレントDL統合検証..."

    saveFileName = "IpkBatch_Report_" & Format(Now, "HHmmss") & ".csv"
    fullSavePath = saveDir & "\" & saveFileName
    If Dir(fullSavePath) <> "" Then Kill fullSavePath

    ' 1. 親画面(P01)でバッチ帳票のサイレントダウンロードを事前予約
    rpaEngine.RunAction "Enable-SilentDownload", CreateParams("DownloadDirectory", saveDir, "FileName", saveFileName)

    ' 2. オーバーライト検証用関数 Set-WebTextInput_Custom で開始日を更新（#dsp_ を検知して自動分岐）
    rpaEngine.RunAction "Set-WebTextInput_Custom", CreateParams( _
        "Selector", "#dsp_IP999SC\~tfdKaisiBi", _
        "Value", "5080401")

    QueryPerformanceCounter startTime
    ' 3. ピッカーが取得した #HJ-SET-FOCUSFUNCTION11 をそのまま Invoke-WebClick_Custom に渡してクリック
    '    （内部で #HJ-SET-FOCUSFUNCTION11 にフォーカス移動後、親の #FUNC11 をクリックしてバッチDL発火）
    rpaEngine.RunAction "Invoke-WebClick_Custom", CreateParams("Selector", "#HJ-SET-FOCUSFUNCTION11")

    ' 4. 仮想タブ側で発火したバッチ帳票CSVダウンロードの完了を待機
    result = rpaEngine.RunAction("Wait-FileDownload", CreateParams("FilePath", fullSavePath, "TimeoutSec", 30))
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/F11バッチ仮想タブDL (TimeMs) : " & Format(elapsedTimeMs, "0.00")

    If Dir(fullSavePath) <> "" Then
        Debug.Print " ◆ [OK] 保存/F11バッチ帳票CSV: " & fullSavePath
    Else
        Err.Raise 999, "Test", "F11バッチ帳票CSVファイルが見つかりません"
    End If

    ' 5. 子ウィンドウの自動クローズと抽出期間反映ステータスの確認
    Sleep 1200
    Debug.Print " ◆ [比較後] F11バッチ送信結果       : " & rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#f11_result"))
    Debug.Print " ◆ [比較後] 親画面(P01)ステータス   : " & rpaEngine.RunAction("Get-WebText", CreateParams("Selector", "#statusMsg"))
    Debug.Print " ◆ [比較後] 自動クローズ後のタブ一覧: " & rpaEngine.RunAction("List-Tabs")
    
    TEST_Part23 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名)
    Call HandleError_Type1("TEST_Part23")
    TEST_Part23 = False
End Function

Private Function TEST_Part99() As Boolean
    Debug.Print "=========================================================="
    Debug.Print "Part 99: "
    Debug.Print "=========================================================="
    On Error GoTo ErrorHandler
''' xxx
    
    TEST_Part99 = True
    Exit Function
ErrorHandler:
    ' --- 共通エラー処理(現在のプロシージャ名)
    Call HandleError_Type1("TEST_Part99")
    TEST_Part99 = False
End Function

