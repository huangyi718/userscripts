// ==UserScript==
// @name         FIF阅卷助手-开源版
// @namespace    http://tampermonkey.net/
// @version      5.3
// @description  A键映射到系统原生的+0按钮，支持全按键自定义，带冲突检测。
// @author       HY
// @match        *://aimark.fifedu.com/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @run-at       document-end
// @updateURL    https://raw.githubusercontent.com/huangyi718/userscripts/main/fif-helper.user.js
// @downloadURL  https://raw.githubusercontent.com/huangyi718/userscripts/main/fif-helper.user.js
// ==/UserScript==

(function() {
    'use strict';

    // 1. 初始化配置
    const defaultKeys = {
        right: 'q', halfM: 'w', wrong: 'e',
        score0: 'a', // 映射到 +0
        dotFive: '`', // 映射到 +0.5
        s1: '1', s2: '2', s3: '3', s4: '4', s5: '5'
    };
    let userKeys = GM_getValue('fif_custom_keys_v51', defaultKeys);

    // 2. 深度寻找元素
    function findTargetRecursive(doc, selector) {
        let el = doc.querySelector(selector);
        if (el) return el;
        const iframes = document.querySelectorAll('iframe');
        for (let f of iframes) {
            try {
                el = findTargetRecursive(f.contentWindow.document, selector);
                if (el) return el;
            } catch (e) {}
        }
        return null;
    }

    // 3. 模拟点击
    function triggerClick(el) {
        if (!el) return;
        const ownerWindow = el.ownerDocument.defaultView || window;
        try {
            const events = ['mousedown', 'mouseup', 'click'];
            events.forEach(type => {
                const ev = new MouseEvent(type, { view: ownerWindow, bubbles: true, cancelable: true, buttons: 1 });
                el.dispatchEvent(ev);
            });
        } catch (err) { el.click(); }
    }

    function ensureSpotMarkingOn() {
        const spotBtn = findTargetRecursive(document, '#spotMarking');
        if (spotBtn && !spotBtn.classList.contains('on') && !spotBtn.classList.contains('active')) {
            triggerClick(spotBtn);
        }
    }

    // 4. 核心动作逻辑
    function setScore(val, isHalf = false) {
        ensureSpotMarkingOn();
        const selector = isHalf
            ? `a[name="actionTypeBtnHalf"][spotscore="0.5"]`
            : `a[name="actionTypeBtn"][spotscore="${val}"]`;

        const btn = findTargetRecursive(document, selector);
        if (btn) {
            triggerClick(btn);
            if (!isHalf) {
                setTimeout(() => {
                    const btnHalf = findTargetRecursive(document, `a[name="actionTypeBtnHalf"][spotscore="0.5"]`);
                    if (btnHalf && btnHalf.classList.contains('on')) triggerClick(btnHalf);
                }, 30);
            }
        }
    }

    // 5. 键盘监听
    document.addEventListener('keydown', function(e) {
        if (['INPUT', 'TEXTAREA'].includes(e.target.tagName) || e.target.isContentEditable) return;
        const key = e.key.toLowerCase();

        if (key === userKeys.score0) { setScore('0'); return; }
        if (key === userKeys.dotFive || (userKeys.dotFive === '`' && e.code === 'Backquote')) { setScore('0.5', true); return; }

        const scoreMap = { [userKeys.s1]:'1', [userKeys.s2]:'2', [userKeys.s3]:'3', [userKeys.s4]:'4', [userKeys.s5]:'5' };
        if (scoreMap[key]) { setScore(scoreMap[key]); return; }

        if (key === userKeys.right) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="right"]`));
        if (key === userKeys.halfM) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="half"]`));
        if (key === userKeys.wrong) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="wrong"]`));
    }, true);

    // 6. UI 设置面板
    GM_addStyle(`
        #fif-cfg-btn { position: fixed; top: 20%; right: 0; z-index: 9999; background: #4CAF50; color: #fff; padding: 10px 5px; cursor: pointer; writing-mode: vertical-lr; border-radius: 5px 0 0 5px; font-size: 12px; box-shadow: -2px 2px 5px rgba(0,0,0,0.2); }
        #fif-panel { position: fixed; top: 20%; right: 40px; z-index: 9999; background: #fff; border: 1px solid #ddd; padding: 15px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); display: none; width: 220px; font-family: sans-serif; }
        .row { margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; color: #333; }
        .row input { width: 45px; text-align: center; border: 1px solid #ccc; border-radius: 3px; padding: 2px; transition: border 0.3s; }
        .row input:focus { border-color: #4CAF50; outline: none; }
        .row input.readonly { background: #f5f5f5; color: #888; cursor: not-allowed; }
        .tip { font-size: 11px; color: #666; margin-bottom: 10px; line-height: 1.4; border-left: 3px solid #ff9800; padding-left: 5px; }
        #save-btn { width: 100%; padding: 8px; background: #4CAF50; color: #fff; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        #save-btn:hover { background: #45a049; }
    `);

    const panel = document.createElement('div');
    panel.id = 'fif-panel';
    panel.innerHTML = `
        <div style="font-weight:bold; margin-bottom:10px; border-bottom:1px solid #eee; padding-bottom:5px;">按键配置 (v5.3)</div>
        <div class="tip">提示：数字键建议保持默认，避免与其他快捷键冲突。</div>
        <div class="row">对/半/错: <div>
            <input type="text" id="r" value="${userKeys.right}">
            <input type="text" id="h" value="${userKeys.halfM}">
            <input type="text" id="w" value="${userKeys.wrong}">
        </div></div>
        <div class="row">0 分按键 (归零): <input type="text" id="s0" value="${userKeys.score0}"></div>
        <div class="row">0.5 辅助键: <input type="text" id="sf" value="${userKeys.dotFive}"></div>
        <div class="row" style="color:#999">1-5 分按键 (不建议改): <div>
            <input type="text" id="s1" class="readonly" readonly value="${userKeys.s1}">
            <input type="text" id="s2" class="readonly" readonly value="${userKeys.s2}">
            <input type="text" id="s3" class="readonly" readonly value="${userKeys.s3}">
            <input type="text" id="s4" class="readonly" readonly value="${userKeys.s4}">
            <input type="text" id="s5" class="readonly" readonly value="${userKeys.s5}">
        </div></div>
        <button id="save-btn">保存配置并刷新</button>
    `;
    document.body.appendChild(panel);

    const btn = document.createElement('div');
    btn.id = 'fif-cfg-btn';
    btn.innerText = '快捷键设置';
    btn.onclick = () => { panel.style.display = panel.style.display === 'none' ? 'block' : 'none'; };
    document.body.appendChild(btn);

    document.getElementById('save-btn').onclick = () => {
        const idMap = { r:'right', h:'halfM', w:'wrong', s0:'score0', sf:'dotFive', s1:'s1', s2:'s2', s3:'s3', s4:'s4', s5:'s5' };
        const newKeys = {};
        const usedValues = new Set();
        let conflict = false;

        // 收集并检测冲突
        for (let id in idMap) {
            const val = document.getElementById(id).value.toLowerCase().trim();
            if (!val) {
                alert("按键不能为空！");
                return;
            }
            if (usedValues.has(val)) {
                alert(`按键冲突：字符 "${val}" 被重复使用了，请重新设置！`);
                conflict = true;
                break;
            }
            usedValues.add(val);
            newKeys[idMap[id]] = val;
        }

        if (!conflict) {
            userKeys = newKeys;
            GM_setValue('fif_custom_keys_v51', userKeys);
            alert("配置保存成功，页面即将刷新。");
            location.reload();
        }
    };
})();
