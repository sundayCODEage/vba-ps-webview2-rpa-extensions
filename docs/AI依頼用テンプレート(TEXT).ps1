あなたは「VBA × PowerShell 5.1 × WebView2 汎用RPAエンジン (v303)」の拡張モジュール開発エキスパートです。
基幹エンジン（`Ps_Engine_Core_v303.ps1` 等）は完成・正常稼働しています。
以下の【エンジン前提仕様・規約】と【標準コードテンプレート】を厳守し、私が後述する「対象HTML構造とやりたい操作」に合わせた拡張関数（PowerShellコード）を作成してください。

### 【エンジン前提仕様・コーディング規約】
1. **動的ディスパッチ規約**:
   - VBAからのJSONコマンドは、Core側で `$res = & $method @parameters | Out-String` として動的実行され、標準出力の `[RESULT]` 経由でVBAへ返却される。
2. **命名・比較演算規約**:
   - PowerShellパラメータ・グローバル変数: `PascalCase`（時間系は単位 `Sec` / `Ms` 必須、真偽値文字列引数は接尾辞 `Flg` 必須）。
   - PowerShellローカル変数・JS埋め込み用Base64変数: `camelCase`（例: `$Selector` ➔ `$selB64`）。
   - JavaScript内部コード: `var` 宣言は原則禁止。`const` / `let` および `camelCase` に統一する。
   - `$null` との等価比較: 配列フィルタリング誤動作を防ぐため、常に左辺に `$null` を配置する（例: `if ($null -ne $res)`）。
3. **JSインジェクション安全基準とラッパー制約**:
   - 動的な文字列をJSヒアドキュメント（`@"..."@`）へ埋め込む際は、必ず `ConvertTo-JsSafeBase64` でエンコードし、JS側で `decodeURIComponent(escape(window.atob('...')))` により復元する。
   - **【重要】JSヒアドキュメントの最終行に `//` 単一行コメントを記述してはならない**（下層ラッパーの閉じカッコ `})();` がコメントアウトされ構文エラーになるため）。
   - JS側の戻り値は型統一ラッパーによりPowerShell側へそのまま型維持で返るため、真偽値は文字列比較（`-eq "true"`）ではなく `[bool]` として直接評価する。
4. **Coreおよびブラウザ提供の共通関数（定義不要・呼び出し専用）**:
   - `[PS] Invoke-WebScript -Js <string> [-Retries <int>]` : CDPまたはNative経由でJSを実行し、結果（型維持またはJSON文字列）を返す。
   - `[PS] ConvertTo-JsSafeBase64 -Text <string>` : 文字列をUTF-8 Base64へ変換する。
   - `[PS] Wait-Condition -ConditionBlock <scriptblock> -TimeoutSec <int> -TimeoutMessage <string> -CallerFunc <string>` : `DoEvents()` を伴うポーリング待機（タイムアウト時は自動で `New-EngineException` をthrow）。
   - `[PS] New-EngineException -Func <string> -Type <Timeout|未発見|JSエラー|CDPエラー|ネイティブエラー|初期化エラー|ファイルエラー|UIAエラー|引数エラー|内部エラー> -Message <string> [-Details <string>]` : VBA連携用の標準例外文字列を生成する。
   - `[PS] Write-DebugLog -Message <string> -Level <Info|Success|Warning|Error|Fatal>` : 実行ログ出力。
   - `[JS] utilFindInFrames(window, function(win) { ... })` : ブラウザ初期化時に全画面へ自動注入済み。同一オリジンの多段 `iframe` / `frame` を再帰探索し、コールバックが `null` / `false` / `undefined` 以外を返した時点でその値を返却する。
   - `[JS] deepQuerySelector(selector, rootDoc)` : 自動注入済み。Shadow DOMの内部を再帰的に貫通して要素を探索する。
   - `[JS] window.DOMUtils` : 自動注入済み（`isVisible(el)`, `generateCssSelector(el, root)`, `safeClick(el)`, `cssEscape(str)`）。
5. **コメント規約**:
   - 関数見出しは `# --- 見出し（体言止め） ---` とし、次行に `# 補足説明（〜する。）` を常体で記述する。関数内の個別コメントに `[ ]` は使用しない。

### 【標準コードテンプレート（新規拡張関数 ＋ 既存関数オーバーライト）】
```powershell
# --- カスタム要素操作テンプレート ---
# 多段iframeおよびShadow DOMを透過探索し、複合UI要素の操作と結果JSONの返却を行う。
function Invoke-WebCustomTemplate {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "[$func] 開始: Selector='$Selector', Value='$Value'" -Level Info

    # JS埋め込み用パラメータのBase64化： 構文破壊を防ぐためUTF-8 Base64へ変換する。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value

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
        $rawRes = Invoke-WebScript -Js $jsCode
        if (-not [string]::IsNullOrEmpty($rawRes)) {
            try {
                $parsed = $rawRes | ConvertFrom-Json
                if ($parsed.status -eq "OK") {
                    $script:customRes = $rawRes
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
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )
    if ($Selector -match '^#?custom_prefix_') {
        return Invoke-WebCustomTemplate -Selector $Selector -Value $Value -TimeoutSec $TimeoutSec
    }
    return Original_Set-WebTextInput @PSBoundParameters
}
