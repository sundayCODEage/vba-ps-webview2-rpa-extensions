# ------------------------------------------------------------------------------#●#✘
<#
  [専用拡張モジュール]：2層構造日付タグ（DateTag）・汎用ペア入力の制御・フォーカス連動ファンクションキー
#>
# ------------------------------------------------------------------------------
<#
    ============================================================================
   [1. 専用拡張コマンド]
   - Set-WebIpkDateInput     : 2層構造の和暦日付コンポーネント（dsp_表示用 + hidden送信用）への入力・確定
   - Set-WebPairedInput      : 表示用要素と送信用hidden要素のペア更新および仮想DOM（React/Vue等）のState同期
   - Invoke-WebIpkFuncClick  : フォーカス受け要素を経由したファンクションボタン（#FUNCxx）のクリック

■ [2. ドットソース・オーバーライト検証用ラッパー関数] ※テスト用コード
   - Set-WebTextInput_Custom : 対象要素が #dsp_ 等の複合UIなら専用処理へ分岐し、通常要素なら標準関数へ委譲する
   - Invoke-WebClick_Custom  : 対象要素が #HJ-SET-FOCUSFUNCTIONxx / #FUNCxx なら専用処理へ分岐し、通常要素なら標準関数へ委譲する
   ※ 手動オーバーライト確認時は、関数名末尾の「_Custom」を削除することで標準関数（Lib-WebAction）を上書きできる。
    ============================================================================
#>

# ------------------------------------------------------------------------------
# [複合UIコンポーネント操作]： 業務システム固有の2層構造UIやイベント依存要素を制御する。

# --- IPK 2層構造日付入力 ---
# IPK財務会計システムのDateTag（dsp_表示用 + hidden送信用）へ7桁和暦を入力し確定させる。
function Set-WebIpkDateInput {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [string]$DisplayValue = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
#●    Write-DebugLog -Message "[$func] 開始: Selector='$Selector', Value='$Value'" -Level Info

    # セレクタの正規化： CSSエスケープや接頭辞（#, dsp_）を除去してベース項目名を抽出する。
    $cleanName = ($Selector -replace '\\~', '~').TrimStart('#')
    if ($cleanName.StartsWith("dsp_")) {
        $cleanName = $cleanName.Substring(4)
    }
    $dspId   = "dsp_$cleanName"
    $hidName = $cleanName

    # 和暦表示文字列の自動生成： DisplayValue省略時、7桁和暦（GYYMMDD）から表示文字列を算出する。
    $dispVal = $DisplayValue
    if ([string]::IsNullOrEmpty($dispVal) -and $Value -match '^([1-5])(\d{2})(\d{2})(\d{2})$') {
        $gengoMap = @{
            "5" = "令和"
            "4" = "平成"
            "3" = "昭和"
            "2" = "大正"
            "1" = "明治"
        }
        $dispVal = "{0}{1}年{2}月{3}日" -f $gengoMap[$Matches[1]], $Matches[2], $Matches[3], $Matches[4]
    }

    # JS埋め込み用パラメータのBase64化： 構文破壊を防ぐためUTF-8 Base64へ変換する。
    $dspIdB64   = ConvertTo-JsSafeBase64 -Text $dspId
    $hidNameB64 = ConvertTo-JsSafeBase64 -Text $hidName
    $val7B64    = ConvertTo-JsSafeBase64 -Text $Value
    $dispB64    = ConvertTo-JsSafeBase64 -Text $dispVal

    $jsCode = @"
return (function() {
    const fromB64 = (s) => s ? decodeURIComponent(escape(window.atob(s))) : '';
    const dspId   = fromB64('$dspIdB64');
    const hidName = fromB64('$hidNameB64');
    const val7    = fromB64('$val7B64');
    const dispStr = fromB64('$dispB64');

    // トップ画面および同一オリジンframe/iframe内を再帰探索する。
    function findIpkElements(doc) {
        let dspEl = doc.getElementById(dspId);
        let hidEl = null;
        const inputs = doc.getElementsByTagName('input');
        for (let i = 0; i < inputs.length; i++) {
            if (inputs[i].type === 'hidden' && (inputs[i].name === hidName || inputs[i].id === hidName)) {
                hidEl = inputs[i];
                break;
            }
        }
        if (!hidEl && dspEl && dspEl.nextElementSibling && dspEl.nextElementSibling.type === 'hidden') {
            hidEl = dspEl.nextElementSibling;
        }
        if (dspEl || hidEl) {
            return { dsp: dspEl, hid: hidEl, win: doc.defaultView || window };
        }
        const frames = doc.querySelectorAll('iframe, frame');
        for (let i = 0; i < frames.length; i++) {
            try {
                const subDoc = frames[i].contentDocument || (frames[i].contentWindow && frames[i].contentWindow.document);
                if (subDoc) {
                    const found = findIpkElements(subDoc);
                    if (found) return found;
                }
            } catch (e) {}
        }
        return null;
    }

    const target = findIpkElements(document);
    if (!target || (!target.dsp && !target.hid)) {
        return JSON.stringify({ status: 'NOT_FOUND', dspId: dspId });
    }

    const dsp = target.dsp;
    const hid = target.hid;
    const win = target.win;

    if (dsp) {
        // フォーカス処理の発火： ネイティブfocusを優先し、未変換時のみcustomOnFocusを呼び出す。
        try {
            dsp.focus();
            if ((!hid || dsp.value !== hid.value) && win.DateTag && typeof win.DateTag.customOnFocus === 'function') {
                win.DateTag.customOnFocus(dsp);
            } else if ((!hid || dsp.value !== hid.value) && typeof dsp.onfocus === 'function') {
                dsp.onfocus.call(dsp, new win.FocusEvent('focus'));
            }
            if (win.DateTag && typeof win.DateTag.customOnMouseup === 'function') {
                win.DateTag.customOnMouseup(dsp);
            }
        } catch (e) {}

        // 7桁和暦数値の設定： 入力値をセットし各種変更イベントを発火させる。
        dsp.value = val7;
        dsp.dispatchEvent(new win.Event('input',  { bubbles: true }));
        dsp.dispatchEvent(new win.Event('change', { bubbles: true }));

        // フォーカスアウト処理の発火： ネイティブblurを優先し、7桁のまま未変換の場合のみcustomOnblurを呼び出す。
        try {
            dsp.blur();
            if (dsp.value === val7) {
                const blurEvt = new win.FocusEvent('blur', { bubbles: false, cancelable: true });
                if (win.DateTag && typeof win.DateTag.customOnblur === 'function') {
                    win.DateTag.customOnblur(blurEvt, dsp, '#FFE6ED');
                } else if (typeof dsp.onblur === 'function') {
                    dsp.onblur.call(dsp, blurEvt);
                }
            }
        } catch (e) {}

        // 表示値・orgvalue・missdate属性の同期： 非アクティブ時でも表示値と保持属性を一致させる。
        if (dispStr !== '') {
            dsp.value = dispStr;
            dsp.setAttribute('orgvalue', dispStr);
        } else {
            dsp.setAttribute('orgvalue', dsp.value);
        }
        if (dsp.hasAttribute('missdate')) {
            dsp.setAttribute('missdate', 'false');
        }
    }

    // 送信用hidden要素の更新： サーバー送信用の実データ値を更新する。
    if (hid) {
        hid.value = val7;
    }

    return JSON.stringify({
        status:   'OK',
        dspValue: dsp ? dsp.value : '',
        orgValue: dsp ? dsp.getAttribute('orgvalue') : '',
        hidName:  hid ? (hid.name || hid.id) : '(none)',
        hidValue: hid ? hid.value : ''
    });
})();
"@

    $script:ipkDateRes = ""

    # 要素出現および更新完了をポーリング待機する。
    Wait-Condition -TimeoutSec $TimeoutSec -TimeoutMessage "対象の日付要素が見つかりませんでした (dspId=$dspId)" -CallerFunc $func -ConditionBlock {
        $rawRes = Invoke-WebScript -Js $jsCode
        if (-not [string]::IsNullOrEmpty($rawRes)) {
            try {
                $parsed = $rawRes | ConvertFrom-Json
                if ($parsed.status -eq "OK") {
                    $script:ipkDateRes = $rawRes
                    Write-DebugLog -Message "[$func] 成功: dsp='$($parsed.dspValue)' (org='$($parsed.orgValue)'), hidden($($parsed.hidName))='$($parsed.hidValue)'" -Level Success
                    return $true
                }
            } catch {}
        }
        return $false
    } | Out-Null

    return $script:ipkDateRes
}

# --- 汎用ペア要素入力 ---
# 表示用要素と送信用hidden要素のペア更新、およびReact/Vue等の仮想DOM State更新を一括実行する。
function Set-WebPairedInput {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [string]$DisplayValue = "",
        [string]$HiddenSelector = "",
        [Alias("ForceReadonly")][string]$ForceReadonlyFlg = "true",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
#●    Write-DebugLog -Message "[$func] 開始: Selector='$Selector', Value='$Value', Hidden='$HiddenSelector'" -Level Info

    # 真偽値フラグの変換： 文字列引数をBooleanへ変換し、JS埋め込み用の真偽値文字列を生成する。
    $forceRo = try { [System.Convert]::ToBoolean($ForceReadonlyFlg) } catch { $true }
    $forceRoJs = if ($forceRo) { "true" } else { "false" }

    # JS埋め込み用パラメータのBase64化： 各引数をUTF-8 Base64へ変換する。
    $selB64    = ConvertTo-JsSafeBase64 -Text $Selector
    $hidSelB64 = ConvertTo-JsSafeBase64 -Text $HiddenSelector
    $valB64    = ConvertTo-JsSafeBase64 -Text $Value
    $dispB64   = ConvertTo-JsSafeBase64 -Text $DisplayValue

    $jsCode = @"
return (function() {
    const fromB64 = (s) => s ? decodeURIComponent(escape(window.atob(s))) : '';
    const sel     = fromB64('$selB64');
    const hidSel  = fromB64('$hidSelB64');
    const rawVal  = fromB64('$valB64');
    const dispVal = fromB64('$dispB64');
    const forceRo = $forceRoJs;

    // Shadow DOMおよび同一オリジンframe/iframeを透過して要素を探索する。
    function findEl(doc, query) {
        if (!query) return null;
        try {
            if (doc.defaultView && typeof doc.defaultView.deepQuerySelector === 'function') {
                const el = doc.defaultView.deepQuerySelector(query, doc);
                if (el) return el;
            }
            const direct = doc.querySelector(query);
            if (direct) return direct;
        } catch (e) {}

        const frames = doc.querySelectorAll('iframe, frame');
        for (let i = 0; i < frames.length; i++) {
            try {
                const subDoc = frames[i].contentDocument || (frames[i].contentWindow && frames[i].contentWindow.document);
                if (subDoc) {
                    const found = findEl(subDoc, query);
                    if (found) return found;
                }
            } catch (e) {}
        }
        return null;
    }

    const dsp = findEl(document, sel);
    if (!dsp) {
        return JSON.stringify({ status: 'NOT_FOUND', selector: sel });
    }

    const win = (dsp.ownerDocument || document).defaultView || window;

    // 送信用hidden要素の特定： 明示指定 -> 隣接要素 -> 同一親要素内の順で探索する。
    let hid = hidSel ? findEl(document, hidSel) : null;
    if (!hid) {
        if (dsp.nextElementSibling && dsp.nextElementSibling.type === 'hidden') {
            hid = dsp.nextElementSibling;
        } else if (dsp.parentElement) {
            hid = dsp.parentElement.querySelector('input[type="hidden"]');
        }
    }

    // 仮想DOM（React/Vue等）のStateトラッカーを反応させるためプロトタイプのネイティブセッターを呼び出す。
    function setNativeValue(el, val) {
        try {
            const desc1 = Object.getOwnPropertyDescriptor(el, 'value');
            const desc2 = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value');
            const setter = desc1 ? desc1.set : null;
            const protoSetter = desc2 ? desc2.set : null;
            if (setter && setter !== protoSetter && protoSetter) {
                protoSetter.call(el, val);
            } else if (setter) {
                setter.call(el, val);
            } else {
                el.value = val;
            }
        } catch (e) {
            el.value = val;
        }
    }

    // readonly属性の一時解除： カレンダー入力欄等の書き込み制限を一時的に解除する。
    const wasRo = dsp.readOnly;
    if (forceRo && wasRo) {
        dsp.readOnly = false;
    }

    // フォーカス・値変更・フォーカスアウトの一連のイベントを発火させる。
    try {
        dsp.focus();
        dsp.dispatchEvent(new win.FocusEvent('focus', { bubbles: false }));
    } catch (e) {}

    setNativeValue(dsp, rawVal);
    dsp.dispatchEvent(new win.Event('input',  { bubbles: true }));
    dsp.dispatchEvent(new win.Event('change', { bubbles: true }));

    try {
        dsp.blur();
        dsp.dispatchEvent(new win.FocusEvent('blur', { bubbles: false }));
    } catch (e) {}

    // 表示用文字列と保持属性の同期： DisplayValue指定時は表示値とorgvalue等を更新する。
    if (dispVal !== '') {
        setNativeValue(dsp, dispVal);
        if (dsp.hasAttribute('orgvalue')) {
            dsp.setAttribute('orgvalue', dispVal);
        }
        if (dsp.hasAttribute('data-old-value')) {
            dsp.setAttribute('data-old-value', dispVal);
        }
    }

    if (forceRo && wasRo) {
        dsp.readOnly = true;
    }

    // ペアとなる送信用hidden要素の値を更新する。
    if (hid) {
        setNativeValue(hid, rawVal);
        hid.dispatchEvent(new win.Event('change', { bubbles: true }));
    }

    return JSON.stringify({
        status:   'OK',
        dspValue: dsp.value,
        orgValue: dsp.getAttribute('orgvalue') || '',
        hidName:  hid ? (hid.name || hid.id) : '(none)',
        hidValue: hid ? hid.value : ''
    });
})();
"@

    $script:pairedRes = ""

    # 要素出現および更新完了をポーリング待機する。
    Wait-Condition -TimeoutSec $TimeoutSec -TimeoutMessage "対象要素が見つかりませんでした (Selector=$Selector)" -CallerFunc $func -ConditionBlock {
        $rawRes = Invoke-WebScript -Js $jsCode
        if (-not [string]::IsNullOrEmpty($rawRes)) {
            try {
                $parsed = $rawRes | ConvertFrom-Json
                if ($parsed.status -eq "OK") {
                    $script:pairedRes = $rawRes
                    Write-DebugLog -Message "[$func] 成功: dsp='$($parsed.dspValue)', hidden($($parsed.hidName))='$($parsed.hidValue)'" -Level Success
                    return $true
                }
            } catch {}
        }
        return $false
    } | Out-Null

    return $script:pairedRes
}

# --- IPK ファンクションボタンクリック ---
# フォーカス受けボタン（#HJ-SET-FOCUSFUNCTIONxx）へフォーカスを移動し、前項目のblurを確定させてから親要素（#FUNCxx）をクリックする。
function Invoke-WebIpkFuncClick {
    param (
        [int]$FuncNumber = 0,
        [string]$Selector = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # 対象IDの決定： FuncNumberまたはSelector（#HJ-SET-FOCUSFUNCTION11等も含む）から番号を抽出する。
    $numPart = if ($FuncNumber -gt 0) { "$FuncNumber" } else { $Selector -replace '[^0-9]', '' }
    $funcId  = if (-not [string]::IsNullOrEmpty($numPart)) { "FUNC$numPart" } else { $Selector.TrimStart('#') }
    $focusId = "HJ-SET-FOCUSFUNCTION$numPart"

#●    Write-DebugLog -Message "[$func] 実行: FuncId='#$funcId', FocusId='#$focusId'" -Level Info

    $funcIdB64  = ConvertTo-JsSafeBase64 -Text $funcId
    $focusIdB64 = ConvertTo-JsSafeBase64 -Text $focusId

    $jsCode = @"
return (function() {
    const fromB64 = (s) => s ? decodeURIComponent(escape(window.atob(s))) : '';
    const funcId  = fromB64('$funcIdB64');
    const focusId = fromB64('$focusIdB64');

    // トップ画面および同一オリジンframe/iframe内を再帰探索する。
    function findInDoc(doc) {
        const fEl = doc.getElementById(funcId);
        if (fEl) {
            return {
                funcEl:   fEl,
                focusBtn: doc.getElementById(focusId),
                doc:      doc
            };
        }
        const frames = doc.querySelectorAll('iframe, frame');
        for (let i = 0; i < frames.length; i++) {
            try {
                const subDoc = frames[i].contentDocument || (frames[i].contentWindow && frames[i].contentWindow.document);
                if (subDoc) {
                    const found = findInDoc(subDoc);
                    if (found) return found;
                }
            } catch (e) {}
        }
        return null;
    }

    const target = findInDoc(document);
    if (!target || !target.funcEl) return 'NOT_FOUND';

    // 直前に入力していた項目のonblurを確実に発火させるため、フォーカス受けボタンへフォーカスを移す。
    try {
        if (target.focusBtn) {
            target.focusBtn.focus();
        } else if (target.doc.activeElement && typeof target.doc.activeElement.blur === 'function') {
            target.doc.activeElement.blur();
        }
    } catch (e) {}

    // onclickイベントを持つ親要素をクリックする。
    target.funcEl.click();
    return 'OK';
})();
"@

    Wait-Condition -TimeoutSec $TimeoutSec -TimeoutMessage "対象のファンクション要素 (#$funcId) が見つかりませんでした" -CallerFunc $func -ConditionBlock {
        $res = Invoke-WebScript -Js $jsCode
        return ($res -eq "OK")
    } | Out-Null

    Write-DebugLog -Message "[$func] 成功: #$funcId をクリックしました。" -Level Success
    return "OK"
}

# ==============================================================================
# [ドットソースによる既存関数のオーバーライト機構（比較テスト＆手動切替対応）]
# ==============================================================================

# 標準モジュール（Lib-WebAction_v303.ps1）の元関数を別名で退避する。
if ((Test-Path Function:\Set-WebTextInput) -and -not (Test-Path Function:\Original_Set-WebTextInput)) {
    Copy-Item Function:\Set-WebTextInput Function:\Original_Set-WebTextInput
}
if ((Test-Path Function:\Invoke-WebClick) -and -not (Test-Path Function:\Original_Invoke-WebClick)) {
    Copy-Item Function:\Invoke-WebClick Function:\Original_Invoke-WebClick
}

# --- オーバーライト検証用テキスト入力 ---
# ※ 手動オーバーライト確認時は、関数名末尾の「_Custom」を削除して「Set-WebTextInput」とする。
function Set-WebTextInput_Custom {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [string]$OuterHtml = "",
        [string]$BoundingBox = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    # セレクタが #dsp_ を含む2層構造項目であれば、自動的にカスタム日付入力へルーティングする。
    if ($Selector -match '^#?dsp_') {
        Write-DebugLog -Message "[OverwriteRouting] #dsp_ 要素を検知したため Set-WebIpkDateInput へ自動委譲します ($Selector)" -Level Info
        return Set-WebIpkDateInput -Selector $Selector -Value $Value -TimeoutSec $TimeoutSec
    }

    # 通常要素の場合は退避した標準の Set-WebTextInput をそのまま実行する。
    return Original_Set-WebTextInput @PSBoundParameters
}

# --- オーバーライト検証用クリック ---
# ※ 手動オーバーライト確認時は、関数名末尾の「_Custom」を削除して「Invoke-WebClick」とする。
function Invoke-WebClick_Custom {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [string]$OuterHtml = "",
        [string]$BoundingBox = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    # ピッカーが取得した #HJ-SET-FOCUSFUNCTIONxx または #FUNCxx であれば、Invoke-WebIpkFuncClick へ自動委譲する。
    if ($Selector -match '^#?(HJ-SET-FOCUSFUNCTION|FUNC)\d+$') {
        Write-DebugLog -Message "[OverwriteRouting] ファンクション要素を検知したため Invoke-WebIpkFuncClick へ自動委譲します ($Selector)" -Level Info
        return Invoke-WebIpkFuncClick -Selector $Selector -TimeoutSec $TimeoutSec
    }

    # 通常要素の場合は退避した標準の Invoke-WebClick をそのまま実行する。
    return Original_Invoke-WebClick @PSBoundParameters
}
