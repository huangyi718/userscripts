// ==UserScript==
// @name         FIF阅卷助手-开源版
// @namespace    http://tampermonkey.net/
// @version      5.2
// @description  A键映射到系统原生的+0按钮，支持全按键自定义。
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

            // 如果点的是整数键(包括0)，需要确保 0.5 是熄灭的（除非你后续手动补按~）
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

        // A键 -> 映射到系统的 spotscore="0"
        if (key === userKeys.score0) {
            setScore('0');
            return;
        }

        // ~键 -> 映射到系统的 spotscore="0.5"
        if (key === userKeys.dotFive || (userKeys.dotFive === '`' && e.code === 'Backquote')) {
            setScore('0.5', true);
            return;
        }

        // 1-5 整数
        const scoreMap = { [userKeys.s1]:'1', [userKeys.s2]:'2', [userKeys.s3]:'3', [userKeys.s4]:'4', [userKeys.s5]:'5' };
        if (scoreMap[key]) {
            setScore(scoreMap[key]);
            return;
        }

        // 标注 (QWE)
        if (key === userKeys.right) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="right"]`));
        if (key === userKeys.halfM) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="half"]`));
        if (key === userKeys.wrong) triggerClick(findTargetRecursive(document, `#div_Annotation a[name="AnnotationType"][type="wrong"]`));
    }, true);

    // 6. UI 设置面板
    GM_addStyle(`
        #fif-cfg-btn { position: fixed; top: 20%; right: 0; z-index: 9999; background: #4CAF50; color: #fff; padding: 10px 5px; cursor: pointer; writing-mode: vertical-lr; border-radius: 5px 0 0 5px; font-size: 12px; }
        #fif-panel { position: fixed; top: 20%; right: 40px; z-index: 9999; background: #fff; border: 1px solid #ddd; padding: 15px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.1); display: none; width: 200px; }
        .row { margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; }
        .row input { width: 50px; text-align: center; border: 1px solid #ccc; border-radius: 3px; }
    `);

    const panel = document.createElement('div');
    panel.id = 'fif-panel';
    panel.innerHTML = `
        <div style="font-weight:bold; margin-bottom:10px; border-bottom:1px solid #eee;">自定义按键 (v5.1)</div>
        <div class="row">对/半/错: <div>
            <input type="text" id="r" value="${userKeys.right}">
            <input type="text" id="h" value="${userKeys.halfM}">
            <input type="text" id="w" value="${userKeys.wrong}">
        </div></div>
        <div class="row">0 分按键: <input type="text" id="s0" value="${userKeys.score0}"></div>
        <div class="row">0.5 辅助键: <input type="text" id="sf" value="${userKeys.dotFive}"></div>
        <div class="row">1-5分按键: <div>
            <input type="text" id="s1" style="width:25px" value="${userKeys.s1}">
            <input type="text" id="s2" style="width:25px" value="${userKeys.s2}">
            <input type="text" id="s3" style="width:25px" value="${userKeys.s3}">
            <input type="text" id="s4" style="width:25px" value="${userKeys.s4}">
            <input type="text" id="s5" style="width:25px" value="${userKeys.s5}">
        </div></div>
        <button id="save" style="width:100%; padding:8px; background:#4CAF50; color:#fff; border:none; border-radius:4px; cursor:pointer;">保存配置并刷新</button>
    `;
    document.body.appendChild(panel);

    const btn = document.createElement('div');
    btn.id = 'fif-cfg-btn';
    btn.innerText = '按键设置';
    btn.onclick = () => { panel.style.display = panel.style.display === 'none' ? 'block' : 'none'; };
    document.body.appendChild(btn);

    document.getElementById('save').onclick = () => {
        panel.querySelectorAll('input').forEach(input => {
            const idMap = { r:'right', h:'halfM', w:'wrong', s0:'score0', sf:'dotFive', s1:'s1', s2:'s2', s3:'s3', s4:'s4', s5:'s5' };
            userKeys[idMap[input.id]] = input.value.toLowerCase();
        });
        GM_setValue('fif_custom_keys_v51', userKeys);
        location.reload();
    };
})();
