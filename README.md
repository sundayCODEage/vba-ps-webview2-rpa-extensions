
### 業務システムのUI（2層構造日付・ダミーボタン）を突破する「拡張モジュール＆オーバーライト」(VBA × PowerShell × WebView2)

<h4 align="center">汎用RPA操作エンジン 拡張モジュール＆AI開発テンプレート (Hybrid-Engine v303 Extensions)</h4>

<p align="right">
  <sub>作成日: 2026年11月07日 / 作成者: System beginner & AI Collaborator (Gemini)</sub>
</p>

### 1. はじめに：汎用RPAエンジンと「自分で拡張できるスタイル」
以前、Excel VBAでのIE（DOM）操作によって構築した業務システムRPAが現場では今も稼働していますが、IEのサポート終了や社内PCのセキュリティ制限（外部WebDriverの導入不可、F12開発者ツールの禁止など）を受け、Windows標準の **PowerShell 5.1** と **WebView2ランタイム** だけで動く「汎用RPA操作エンジン (v303)」を開発・公開しました。

<p align="center">
  <img src="https://raw.githubusercontent.com/sundayCODEage/vba-ps-webview2-rpa-engine/main/docs/images/engine_core_anim_Light.svg" width="820" alt="RPA Engine Core Architecture">
</p>

汎用エンジンとして Shadow DOM や 多段iframe、仮想タブのサイレントダウンロードなどに対応した結果、基幹コードは多機能になりましたが、実際にこのエンジンを他の現場で使おうとした時、**「システム特有のクセ（独自UI）に合わせて、特定の機能だけを追加・一部変更したい」** という場面です。

そこで、以下の3つをセットにした 「自分で拡張できるスタイル」を提案します。
1. **プラグイン構造：** 基幹・既存モジュール（安定版）には手を加えず、ドットソースで読み込む（専用）拡張モジュールで新関数の追加や既存関数の上書き（オーバーライト）ができる機構。
2. **ローカル再現テストベッド：** ローカルHTML（`14_popup_dl_and_custom_ui.html`）上で(開発)システムの「エラー」と「解決策」を確認できる環境。
3. <sub>**AI拡張プロンプトテンプレート：**</sub>  _<sub>巨大な基幹コードをすべてAIに読み込ませることなく、最小限のトークンでAIに「エンジンの通信規約と作法」を理解させ、自分の開発システム専用の拡張コードを書かせるためのテンプレート。</sub>_

---

### 2. ベース リポジトリ
エンジン本体のソースコード、導入手順、要素ピッカー（DevTools）等の仕様については、以下のメイン リポジトリにてMITライセンスで公開しています。本リポジトリのコードは、ベース リポジトリのフォルダに追加配置して使用します。

* **メイン リポジトリ（エンジン本体）:**
  [sundayCODEage/vba-ps-webview2-rpa-engine](https://github.com/sundayCODEage/vba-ps-webview2-rpa-engine)

---

### 3. 拡張モジュールの組み込みとドットソース・オーバーライト機構

#### 3.1 ディレクトリ構成とロード順序
**Ps_Engine_Core_v303.ps1** の常時ロード配列（ **$alwaysLoadLibs** ） の **末尾** に、拡張モジュール `Lib-WebCustomUI_v303.ps1` を追加します。

```powershell
# [常時ロード対象モジュールの定義]
$alwaysLoadLibs = @(
    # 基幹モジュール
    "Lib-WebJS_v303.ps1",
    "Lib-WebView2_Init_v303.ps1",
    "Lib-WebView2_Native_v303.ps1",

    # 拡張モジュール
    "Lib-WebAction_v303.ps1",
    "Lib-WebXPath_v303.ps1",
    "Lib-WebDebug_v303.ps1",
    "Lib-DesktopUIA_v303.ps1",
    "Lib-WebSafeAction_v303.ps1",
    "Lib-DevTools_v303.ps1",
    "Lib-WebCustomUI_v303.ps1" # ここ★★
)
```

#### 3.2 ドットソース展開による2つのメリット
1. **動的ディスパッチによる自動認識：** エンジンのメインループは **$res = & $method @parameters | Out-String** でコマンドを動的実行するため、拡張モジュールに関数を追加するだけで、Core側のルーティング定義を変更することなくVBAの RunAction から直接呼び出せます。

2. **既存関数の安全なオーバーライト（条件分岐ラッパー）：**
* PowerShellの **ドットソース**（. $libPath）は後から読み込んだ同名関数で定義を上書きします。そのため、標準モジュール（Lib-WebAction_v303.ps1 等）のコードを直接変更することなく、拡張モジュール側で同名関数を再定義して一時的に挙動を替えることも可能です。<br>
* 元の標準関数を Copy-Item Function:\Set-WebTextInput Function:\Original_Set-WebTextInput で退避した上で同名関数を再定義すれば、「ピッカーで取得した特定の要素（`#dsp_` や `#HJ-SET-FOCUSFUNCTION11`）の時だけカスタム処理へ自動委譲し、通常要素は標準関数へ流す」という透過的なパッチ適用が可能です。

### 4. 業務システム特有の2つの「壁」と解決ロジック (専用拡張: Lib-WebCustomUI_v303.ps1)
業務システム等のイントラネット画面で遭遇した以下の2つの課題UIを、ローカルテストベッド（14_popup_dl_and_custom_ui.html）上に再現して検証しました。

_**課題①：**_ 2層構造の和暦日付コンポーネント (DateTag) の初期値リセット<br>

```html
<input class="date_text" id="dsp_IP999SC~tfdTekiyoBi" value="令和08年11月01日" orgvalue="令和08年11月01日"
       onfocus="DateTag.customOnFocus(this);" onblur="DateTag.customOnblur(arguments[0], this, '#FFE6ED');"
       type="text" maxlength="8" format="GGyy年MM月DD日" missdate="false">
<input type="hidden" name="IP999SC~tfdTekiyoBi" value="5081101">
```

_<sub>画面表示用の `<`input id="dsp_...</sub>_
_<sub>"`>` </sub>_
_<sub>（例：令和08年11月01日）とサーバー送信用の </sub>_
_<sub>`<`input type="hidden"`>`</sub>_
_<sub>（例：5081101）がペアになっており、onfocus </sub>_
_<sub>を通さずに直接値を書き換えると、onblur やF11送信時に orgvalue（初期値）へ強制リセットされる。</sub>_

👉 **（ Set-WebIpkDateInput / Set-WebPairedInput ）：** フレーム内を透過探索し、focus()（7桁数値モード化）➔ 7桁和暦（5080401）セット ➔ blur()（二重発火を防止しつつ和暦整形＆hidden 転写）➔ orgvalue・missdate・hidden の同期を1コマンドで完結させます。

_**課題②：**_ フォーカス受けダミーボタン (#HJ-SET-FOCUSFUNCTION11) の不発<br>

```html
<li id="FUNC11" onclick="Action('F11ACTION',0);document.getElementById('HJ-SET-FOCUSFUNCTION11').focus();">
    <span>F11 </span><p>バッチ起動</p>
    <button type="button" id="HJ-SET-FOCUSFUNCTION11"></button>
</li>
```

_<sub>`<`li id="FUNC11" onclick="Action(...)"`>` の内部にフォーカス制御用の空ボタン `<`button id="HJ-SET-FOCUSFUNCTION11"`>` があり、要素ピッカーが拾った #HJ-SET-FOCUSFUNCTION11 を直接クリックするとイベントが発火しない。</sub>_

👉 **（ Invoke-WebIpkFuncClick ）：** 内部で先に #HJ-SET-FOCUSFUNCTION11 へフォーカスを移して直前の入力欄の onblur を確定させた後、イベント本体を持つ親要素 #FUNC11 をクリックします。

### 5. 統合検証ログ（比較前エラー再現 ➔ 拡張・オーバーライト ➔ 仮想タブCSV出力）
ローカルsandboxにて、①標準関数でのエラー再現、②専用拡張関数、③オーバーライト関数＋仮想タブ経由のバッチ帳票CSVサイレントダウンロードを連続実行した結果です。

```text
==========================================================
Part 23: 業務システム複合UI拡張 ＆ オーバーライト・帳票DL統合テスト
==========================================================
TEST1: >> 標準関数 (Set-WebTextInput / Invoke-WebClick) でのエラー再現検証...
 ◆ [比較前] #HJ-SET-FOCUSFUNCTION11 標準クリック結果: ERROR: #HJ-SET-FOCUSFUNCTION11 直接クリックによりAction不発
 ◆ [比較前] 標準入力後の #FUNC11 送信結果          : NG_RESET: 抽出期間が初期値(5081101)のまま送信されました
 ◆ [比較前] 開始日の表示値（初期値へリセット）     : 令和08年11月01日
==========================================================
TEST2: >> 専用拡張関数による2層構造日付・readonlyペア要素の更新検証...
PS_STDOUT: [LOG] <..> <Success> [Set-WebIpkDateInput] 成功: dsp='令和08年04月01日' (org='令和08年04月01日'), hidden(IP999SC~tfdKaisiBi)='5080401'
 ◆ [拡張関数] 開始日更新JSON : {"status":"OK","dspValue":"令和08年04月01日","orgValue":"令和08年04月01日","hidName":"IP999SC~tfdKaisiBi","hidValue":"5080401"}
PS_STDOUT: [LOG] <..> <Success> [Set-WebPairedInput] 成功: dsp='令和09年03月31日', hidden(IP999SC~tfdSyuryoBi)='5090331'
 ◆ [拡張関数] 終了日更新JSON : {"status":"OK","dspValue":"令和09年03月31日","orgValue":"令和09年03月31日","hidName":"IP999SC~tfdSyuryoBi","hidValue":"5090331"}
==========================================================
TEST3: >> オーバーライト関数 (_Custom) ＋ 仮想タブCSVサイレントDL統合検証...
PS_STDOUT: [LOG] <..> <Info> [OverwriteRouting] #dsp_ 要素を検知したため Set-WebIpkDateInput へ自動委譲します (#dsp_IP999SC\~tfdKaisiBi)
PS_STDOUT: [LOG] <..> <Info> [OverwriteRouting] ファンクション要素を検知したため Invoke-WebIpkFuncClick へ自動委譲します (#HJ-SET-FOCUSFUNCTION11)
PS_STDOUT: [LOG] <..> <Info> [WebView2] 情報: 新規ウィンドウ要求を検知。仮想タブとして捕獲します。
PS_STDOUT: [LOG] <..> <Info> [WebView2] 情報: ポップアップ(6ca593f3-0ea8-4854-a613-ea8474b3a75f)へサイレントダウンロード設定を継承しました。
 ◆ 計測/F11バッチ仮想タブDL (TimeMs) : 1785.93
 ◆ [OK] 保存/F11バッチ帳票CSV: C:\...\RPA_Downloads\IpkBatch_Report_211812.csv
PS_STDOUT: [LOG] <..> <Info> [WebView2] 情報: ゴーストを削除し、親タブ (P01) へ自動復帰しました。
 ◆ [比較後] F11バッチ送信結果       : F11送信完了: Kaisi=5080401 / Syuryo=5090331
 ◆ [比較後] 親画面(P01)ステータス   : 帳票出力完了（親画面復帰）
 ◆ [比較後] 自動クローズ後のタブ一覧: [{"TabId":"P01","Url":".../sandbox/14_popup_dl_and_custom_ui.html","Title":"13. 仮想タブ（ポップアップ）自動DL ＆ 業務複合UI統合検証","IsActive":true}]
```

---
### 6. 自分のシステム開発に合わせてAIで拡張モジュールを作る（AI依頼用テンプレート例）
_<sub>システム開発で特殊なUIに遭遇した際は、`Export-WebHtml` で保存した該当箇所のHTMLタグ等と一緒に、以下のテンプレートをAIへ貼り付けて依頼してみてください。</sub>_

<details>
  <summary>&emsp;📌 <i>( AI依頼用テンプレート例（By GEMINI） ※クリックして展開・コピー )</i></summary>

```powershell
あなたは「VBA × PowerShell 5.1 × WebView2 汎用RPAエンジン (v303)」の拡張モジュール開発エキスパートです。
基幹エンジン（`Ps_Engine_Core_v303.ps1` 等）は完成・正常稼働しています。
以下の【エンジン前提仕様・規約】と【標準コードテンプレート】を厳守し、私が後述する「対象HTML構造とやりたい操作」に合わせた拡張関数（PowerShellコード）を作成してください。

### 【エンジン前提仕様・コーディング規約】
1. コマンド呼び出し規約:
   - VBAからのJSONはCore側で `$res = & $method @parameters | Out-String` として動的実行され、`[RESULT]` としてVBAへ返却される。
2. 命名・比較演算規約:
   - パラメータ・グローバル変数: `PascalCase`（時間系は単位 `Sec`/`Ms` 必須、真偽値文字列引数は接尾辞 `Flg` 必須）。
   - ローカル変数・JS埋め込み用Base64変数: `camelCase`（例: `$Selector` ➔ `$selB64`）。
   - JS内部変数: `var` 禁止、`const`/`let` の `camelCase` に統一する。
   - `$null` 比較: 常に左辺に `$null` を配置する（例: `if ($null -ne $res)`）。
3. JSインジェクション安全基準とラッパー制約:
   - 動的な文字列をJSヒアドキュメント（`@"..."@`）へ埋め込む際は、必ず `ConvertTo-JsSafeBase64` でエンコードし、JS側で `decodeURIComponent(escape(window.atob('...')))` で復元する。
   - 【重要】JSヒアドキュメントの最終行に `//` 単一行コメントを記述してはならない（ラッパーの閉じカッコ `})();` がコメントアウトされ構文エラーになるため）。
   - JS側の戻り値は型統一ラッパーによりそのまま型維持で返るため、真偽値は `[bool]` として直接評価する。
4. Coreおよびブラウザ提供の共通関数（実装不要・呼出専用）:
   - `[PS] Invoke-WebScript -Js <string> [-Retries <int>]` : JSを実行し結果を返す。
   - `[PS] ConvertTo-JsSafeBase64 -Text <string>` : UTF-8のBase64文字列へ変換する。
   - `[PS] Wait-Condition -ConditionBlock <scriptblock> -TimeoutSec <int> -TimeoutMessage <string> -CallerFunc <string>` : `DoEvents()` を伴うポーリング待機（失敗時は自動で `New-EngineException` をthrow）。
   - `[PS] New-EngineException -Func <string> -Type <Timeout|未発見|JSエラー|CDPエラー|ネイティブエラー|初期化エラー|ファイルエラー|UIAエラー|引数エラー|内部エラー> -Message <string> [-Details <string>]` : 標準例外フォーマット生成。
   - `[PS] Write-DebugLog -Message <string> -Level <Info|Success|Warning|Error|Fatal>` : ログ出力。
   - `[JS] utilFindInFrames(window, function(win) { ... })` : 全画面に自動注入済み。同一オリジンの多段 `iframe`/`frame` を再帰探索する。
   - `[JS] deepQuerySelector(selector, rootDoc)` : 自動注入済み。Shadow DOMの内部を再帰的に貫通して探索する。
   - `[JS] window.DOMUtils` : 自動注入済み（`isVisible(el)`, `generateCssSelector(el, root)`, `safeClick(el)`, `cssEscape(str)`）。
5. コメント規約:
   - 関数見出しは `# --- 見出し ---` とし、次行に `# 補足説明（〜する。）` を常体で1行記述する。関数内で `[ ]` は使用しない。

### 【標準コードテンプレート（新規拡張関数 ＋ 既存関数オーバーライト）】

# --- カスタム要素操作テンプレート ---
# 多段iframeおよびShadow DOMを透過探索し、複合UI要素の操作と結果JSONの返却を行う。
function Invoke-WebCustomTemplate {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec =$global:CONFIG.DefaultTimeoutSec
    )

    $func =$MyInvocation.MyCommand.Name
    Write-DebugLog -Message "[$func] 開始: Selector='$Selector', Value='$Value'" -Level Info

    # JS埋め込み用パラメータのBase64化： 構文破壊を防ぐためUTF-8 Base64へ変換する。
    $selB64 = ConvertTo-JsSafeBase64 -Text$Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text$Value

    $jsCode = @"
return (function() {
    const fromB64 = (s) => s ? decodeURIComponent(escape(window.atob(s))) : '';
    const sel = fromB64('$selB64');
    const val = fromB64('$valB64');

    // 自動注入済みの utilFindInFrames と deepQuerySelector で多段iframe・Shadow DOMを透過探索する。
    const target = utilFindInFrames(window, function(win) {
        try {
            const el = (typeof deepQuerySelector === 'function')
                ? deepQuerySelector(sel, win.document)
                : win.document.querySelector(sel);
            if (el) return { el: el, win: win, doc: win.document };
        } catch (e) {}
        return null;
    });

    if (!target) {
        return JSON.stringify({ status: 'NOT_FOUND', selector: sel });
    }

    const el  = target.el;
    const win = target.win;

    // 要素へのカスタム操作とイベント発火を実行する。
    el.focus();
    el.value = val;
    el.dispatchEvent(new win.Event('input',  { bubbles: true }));
    el.dispatchEvent(new win.Event('change', { bubbles: true }));
    el.blur();

    return JSON.stringify({
        status: 'OK',
        resultValue: el.value
    });
})();
"@

    $script:customRes = ""

    # 要素出現および更新完了をポーリング待機する。
    Wait-Condition -TimeoutSec $TimeoutSec -TimeoutMessage "対象要素の操作がタイムアウトしました ($Selector)" -CallerFunc $func -ConditionBlock {
        $rawRes = Invoke-WebScript -Js$jsCode
        if (-not [string]::IsNullOrEmpty($rawRes)) {
            try {
                $parsed =$rawRes | ConvertFrom-Json
                if ($parsed.status -eq "OK") {
                    $script:customRes =$rawRes
                    Write-DebugLog -Message "[$func] 成功: resultValue='$($parsed.resultValue)'" -Level Success
                    return $true
                }
            } catch {}
        }
        return $false
    } | Out-Null

    return $script:customRes
}

# --- 標準関数の条件分岐オーバーライト（必要な場合のみ定義） ---
# 標準関数を Original_* へ退避し、特定セレクタ時のみカスタム関数へ自動委譲する。
if ((Test-Path Function:\Set-WebTextInput) -and -not (Test-Path Function:\Original_Set-WebTextInput)) {
    Copy-Item Function:\Set-WebTextInput Function:\Original_Set-WebTextInput
}
function Set-WebTextInput {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [string]$OuterHtml = "",
        [string]$BoundingBox = "",
        [int]$TimeoutSec =$global:CONFIG.DefaultTimeoutSec
    )
    if ($Selector -match '^#?custom_prefix_') {
        return Invoke-WebCustomTemplate -Selector $Selector -Value $Value -TimeoutSec$TimeoutSec
    }
    return Original_Set-WebTextInput @PSBoundParameters
}
```
</details>

docs/ AI依頼用テンプレート(TEXT).ps1 &emsp; _<sub>構文ハイライト</sub>_

#### 【今回作成してほしい拡張仕様】
- 操作対象のHTMLソース: （ここに画面のHTML抜粋を貼り付ける）　※ サイト、個人情報に注意！
- 実現したい動作・課題: （例：ピッカーで取得すると #xxx が取れるが、実際には親の #yyy をクリックしたい 等）

---
### 📜 ライセンス (License)
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.<br>
このプロジェクトは **MIT License** の下で公開されています。詳しくは [LICENSE](LICENSE) ファイルをご覧ください。

---

Keywords: Excel VBA PowerShell 5.1 WebView2 自作RPA ブラウザ自動化 Selenium不要 WebDriver不要 IEモード移行 2層構造日付 和暦入力 ドットソース 関数オーバーライト サイレントダウンロード AIプロンプト
