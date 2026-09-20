// ==UserScript==
// @name         21tb 重庆专技课一键直达与自动连播
// @namespace    local.codex.cqrl
// @version      4.9.1
// @description  自动进入首个未完成课程、筛选未完成微课、倍速连播，并在全部学完后关闭播放窗口。
// @author       You
// @match        https://*.21tb.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    const TASK_KEY = 'TM_CQRL_AUTO_FLOW_STATE';
    const SPEED_KEY = 'TM_COURSE_PLAY_SPEED';
    const BLOCKED_MICRO_COURSE_TITLES = new Set([
        '车辆应急处理器材、安全防护设施设备管理',
        '突发公共卫生事件现场调查和应急处置（一）'
    ]);
    const PANEL_ID = 'tm-cqrl-helper-panel';
    const PLAYER_PATH = '/els/html/courseStudyItem/courseStudyItem.learn.do';
    const LIST_URL = 'https://cqrl.21tb.com/nms-frontend/index.html#/org/course/list?entrance=zygx';
    const TASK_TTL = 24 * 60 * 60 * 1000;
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    const normalize = (text) => String(text || '').replace(/\s+/g, ' ').trim();

    let navigationBusy = false;
    let autoPlayStarted = false;
    let scoreRefreshScheduled = false;
    const flowChannel = typeof BroadcastChannel === 'function'
        ? new BroadcastChannel('TM_CQRL_AUTO_FLOW_CHANNEL')
        : null;

    function readSpeed() {
        const value = Number.parseFloat(localStorage.getItem(SPEED_KEY));
        return Number.isFinite(value) && value >= 0.5 && value <= 16 ? value : 5;
    }

    function saveSpeed(value) {
        const speed = Math.min(16, Math.max(0.5, Number(value) || 1));
        localStorage.setItem(SPEED_KEY, String(speed));
        return speed;
    }

    function applyVideoSpeed(root = document) {
        const speed = readSpeed();
        root.querySelectorAll('video').forEach(video => {
            try {
                video.muted = true;
                video.volume = 0;
                if (video.playbackRate !== speed) video.playbackRate = speed;
                if (video.paused && !video.ended) video.play().catch(() => {});
            } catch (error) {
                console.debug('[自动学习] 维持视频播放失败：', error);
            }
        });
    }

    // 本脚本会在 iframe 中运行，因此每一层都能维护自己页面里的 video。
    window.setInterval(() => applyVideoSpeed(document), 1000);
    applyVideoSpeed(document);

    // iframe 仅负责倍速；导航、面板和连播调度只允许顶层窗口执行一次。
    if (window.top !== window.self) return;

    function setStatus(text, color = '#f1c40f') {
        const status = document.getElementById('tm-cqrl-status');
        if (status) {
            status.textContent = text;
            status.style.color = color;
        }
    }

    function injectStyles() {
        if (document.getElementById('tm-cqrl-helper-style')) return;
        const style = document.createElement('style');
        style.id = 'tm-cqrl-helper-style';
        style.textContent = `
            #${PANEL_ID} {
                position: fixed; right: 22px; bottom: 24px; z-index: 2147483647;
                min-width: 220px; padding: 13px 15px; border-radius: 10px;
                background: #263747; color: #fff; box-shadow: 0 5px 20px rgba(0,0,0,.28);
                font: 13px/1.55 Arial, "Microsoft YaHei", sans-serif;
            }
            #${PANEL_ID} .tm-title { color: #38d6b3; font-size: 14px; font-weight: 700; margin-bottom: 7px; }
            #${PANEL_ID} .tm-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-top: 8px; }
            #${PANEL_ID} button { width: 100%; padding: 7px 10px; border: 0; border-radius: 18px; color: #fff; background: #1687e8; cursor: pointer; font-weight: 700; }
            #${PANEL_ID} button:hover { background: #0875cf; }
            #${PANEL_ID} input { width: 58px; box-sizing: border-box; padding: 3px 5px; color: #fff; background: #34495e; border: 1px solid #668099; border-radius: 4px; text-align: center; }
            #${PANEL_ID} .tm-note { margin-top: 5px; color: #aebdcc; font-size: 11px; }
        `;
        (document.head || document.documentElement).appendChild(style);
    }

    function createPanel() {
        if (!document.body || document.getElementById(PANEL_ID)) return;
        injectStyles();

        const panel = document.createElement('div');
        panel.id = PANEL_ID;
        panel.innerHTML = `
            <div class="tm-title">重庆专技自动学习助手 v4.9.1</div>
            <div>状态：<span id="tm-cqrl-status">等待操作</span></div>
            <div class="tm-row"><span>播放倍速</span><input id="tm-cqrl-speed" type="number" min="0.5" max="16" step="0.5"></div>
            <div class="tm-row"><button id="tm-cqrl-start" type="button">一键开始全流程</button></div>
            <div class="tm-note">必修、选修目标及剩余学分均从页面自动读取</div>
        `;
        document.body.appendChild(panel);

        const speedInput = panel.querySelector('#tm-cqrl-speed');
        speedInput.value = readSpeed();
        speedInput.addEventListener('change', () => {
            speedInput.value = saveSpeed(speedInput.value);
            applyVideoSpeed(document);
            setStatus(`倍速已设为 ${speedInput.value}x`, '#38d6b3');
        });

        panel.querySelector('#tm-cqrl-start').addEventListener('click', () => {
            writeTask({
                startedAt: Date.now(),
                updatedAt: Date.now(),
                phase: 'must',
                listCategory: '公需科目',
                awaitingPlayer: false,
                listUrl: location.hash.includes('/org/course/list') ? location.href : LIST_URL,
                returnUrl: location.href
            });
            setStatus('全流程已启动…');
            runNavigation();
        });

    }

    function readTask() {
        try {
            const task = JSON.parse(localStorage.getItem(TASK_KEY) || 'null');
            if (!task || Date.now() - Number(task.updatedAt || task.startedAt) > TASK_TTL) {
                localStorage.removeItem(TASK_KEY);
                return null;
            }
            return task;
        } catch (error) {
            localStorage.removeItem(TASK_KEY);
            return null;
        }
    }

    function writeTask(task) {
        task.updatedAt = Date.now();
        localStorage.setItem(TASK_KEY, JSON.stringify(task));
        return task;
    }

    function taskIsActive() {
        return Boolean(readTask());
    }

    function resumeAfterPlayer(finishedAt = Date.now()) {
        const task = readTask();
        if (!task) return;
        const eventTime = Number(finishedAt) || Date.now();
        // postMessage、BroadcastChannel 和直接调用可能同时到达，同一事件只处理一次。
        if (Number(task.lastResumeHandledAt) >= eventTime) return;
        task.lastResumeHandledAt = eventTime;
        task.awaitingPlayer = false;
        task.lastPlayerFinishedAt = eventTime;
        // 给平台留出成绩入库时间；刷新后再从服务端重新取得学分和课程状态。
        task.resumeAfter = Date.now() + 8000;
        writeTask(task);
        setStatus('视频已完成，等待成绩同步并刷新学分…', '#38d6b3');

        const isCourseDetailPage = !location.hash.includes('/org/course/list') &&
            Boolean(document.querySelector('.goods-info'));
        if (isCourseDetailPage && !scoreRefreshScheduled) {
            scoreRefreshScheduled = true;
            window.setTimeout(() => location.reload(), 5000);
        }
    }

    // 暴露一个同源窗口可直接调用的唤醒入口。
    window.__TM_CQRL_RESUME__ = resumeAfterPlayer;

    function isVisible(element) {
        if (!(element instanceof HTMLElement)) return false;
        const rect = element.getBoundingClientRect();
        const style = getComputedStyle(element);
        return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
            style.visibility !== 'hidden' && Number(style.opacity) !== 0;
    }

    function findByExactText(selector, text) {
        return Array.from(document.querySelectorAll(selector)).find(element =>
            isVisible(element) && normalize(element.innerText || element.textContent) === text
        );
    }

    function clickLikeUser(element) {
        if (!element) return false;
        const clickable = element.closest('a, button, [role="button"], .cursor') || element;
        clickable.scrollIntoView({ block: 'center', behavior: 'smooth' });
        try { clickable.focus({ preventScroll: true }); } catch (error) {}
        for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
            clickable.dispatchEvent(new MouseEvent(type, {
                bubbles: true,
                cancelable: true,
                composed: true,
                view: window
            }));
        }
        return true;
    }

    function extractProgress(container) {
        const progressBar = container.querySelector('[role="progressbar"][aria-valuenow]');
        if (progressBar) {
            const value = Number(progressBar.getAttribute('aria-valuenow'));
            if (Number.isFinite(value)) return value;
        }

        const text = normalize(container.innerText || container.textContent);
        const labeled = text.match(/(?:学习)?进度[^%\d]{0,12}(\d+(?:\.\d+)?)\s*%/);
        if (labeled) return Number(labeled[1]);
        const values = [...text.matchAll(/(^|[^\d.])(\d+(?:\.\d+)?)\s*%/g)].map(match => Number(match[2]));
        return values.length === 1 ? values[0] : null;
    }

    function readPageCreditStatus() {
        const goodsInfo = Array.from(document.querySelectorAll('.goods-info')).find(element => {
            const text = normalize(element.innerText || element.textContent);
            return isVisible(element) && text.includes('学分要求') && text.includes('您还需要学习');
        });
        if (!goodsInfo) return null;
        const text = normalize(goodsInfo.innerText || goodsInfo.textContent);
        const remaining = text.match(/您还需要学习\s*(\d+(?:\.\d+)?)\s*必修学分\s*[，,、]?\s*(\d+(?:\.\d+)?)\s*选修学分/);
        if (!remaining) return null;

        const totalRequired = text.match(/学分要求\s*[：:]?\s*(\d+(?:\.\d+)?)/);
        const mustRequired = text.match(/必修\s*[≥>=]+\s*(\d+(?:\.\d+)?)\s*学分/);
        const electiveRequired = text.match(/选修\s*[≥>=]+\s*(\d+(?:\.\d+)?)\s*学分/);
        return {
            totalRequired: totalRequired ? Number(totalRequired[1]) : null,
            mustRequired: mustRequired ? Number(mustRequired[1]) : null,
            electiveRequired: electiveRequired ? Number(electiveRequired[1]) : null,
            mustRemaining: Number(remaining[1]),
            electiveRemaining: Number(remaining[2])
        };
    }

    async function waitForStableCreditStatus(timeout = 18000) {
        // 页面初次渲染时剩余学分可能暂时显示为 0，先等待接口数据回填。
        await sleep(3000);
        const deadline = Date.now() + timeout;
        let latest = null;
        let previousSignature = '';
        let stableCount = 0;

        while (Date.now() < deadline) {
            const loading = Array.from(document.querySelectorAll('.el-loading-mask'))
                .some(element => isVisible(element));
            const current = loading ? null : readPageCreditStatus();

            if (current && current.mustRequired !== null && current.electiveRequired !== null) {
                const signature = JSON.stringify(current);
                stableCount = signature === previousSignature ? stableCount + 1 : 1;
                previousSignature = signature;
                latest = current;

                // 0/0 更可能是加载占位值，因此要求更长的稳定时间。
                const requiredSamples = current.mustRemaining === 0 && current.electiveRemaining === 0 ? 5 : 3;
                if (stableCount >= requiredSamples) return current;
            } else {
                stableCount = 0;
                previousSignature = '';
            }
            await sleep(700);
        }
        return latest;
    }

    function findFirstIncompleteEnterButton() {
        // 总列表真实结构：section.course__item > .course__progress .num + button.enter-btn。
        const courseCards = Array.from(document.querySelectorAll('.course__list--content > section.course__item, section.course__item'))
            .filter(isVisible);
        for (const card of courseCards) {
            const progressText = normalize(card.querySelector('.course__progress .num')?.textContent);
            const match = progressText.match(/(\d+(?:\.\d+)?)\s*%/);
            const progress = match ? Number(match[1]) : extractProgress(card);
            if (progress === null || progress >= 100) continue;

            const button = card.querySelector('.course__btn--outer button.enter-btn, button.enter-btn');
            if (button && isVisible(button)) {
                const title = normalize(card.querySelector('.course__title')?.textContent);
                return { button, progress, card, title };
            }
        }

        // 保留按钮反向查找作为兼容后备。
        const buttons = Array.from(document.querySelectorAll('button.enter-btn, button, a, [role="button"]'))
            .filter(button => isVisible(button) && /进入学习/.test(normalize(button.innerText || button.textContent)));

        for (const button of buttons) {
            let node = button.parentElement;
            let fallback = null;
            for (let depth = 0; node && node !== document.body && depth < 9; depth += 1, node = node.parentElement) {
                const progress = extractProgress(node);
                if (progress === null) continue;
                fallback = { button, progress };
                const count = node.querySelectorAll('button.enter-btn').length;
                if (count <= 1) break;
            }
            if (fallback && fallback.progress < 100) return fallback;
        }
        return null;
    }

    function getProjectCategoryBox() {
        const categoryRow = Array.from(document.querySelectorAll('.form-item')).find(row => {
            const title = row.querySelector('.title');
            return title && /^项目分类[：:]?$/.test(normalize(title.innerText || title.textContent));
        });
        return categoryRow?.querySelector('.item-box') || null;
    }

    function findListCategory(name) {
        const categoryBox = getProjectCategoryBox();
        if (!categoryBox) return null;
        return Array.from(categoryBox.querySelectorAll('.item'))
            .find(element => isVisible(element) && normalize(element.innerText || element.textContent) === name);
    }

    function currentListCategory() {
        const active = getProjectCategoryBox()?.querySelector('.item.active-item');
        return active ? normalize(active.innerText || active.textContent) : '';
    }

    function readPagination() {
        const pagination = document.querySelector('.course__pagination .el-pagination') ||
            Array.from(document.querySelectorAll('.el-pagination')).find(isVisible);
        if (!pagination || !isVisible(pagination)) return null;
        const active = pagination.querySelector('.el-pager .number.active');
        const next = pagination.querySelector('button.btn-next');
        const pageNumbers = Array.from(pagination.querySelectorAll('.el-pager .number'))
            .map(element => Number(normalize(element.textContent)))
            .filter(Number.isFinite);
        return {
            current: active ? Number(normalize(active.textContent)) : 1,
            lastVisible: pageNumbers.length ? Math.max(...pageNumbers) : 1,
            next,
            hasNext: Boolean(next && !next.disabled && !next.hasAttribute('disabled'))
        };
    }

    async function runNavigation() {
        if (navigationBusy || !taskIsActive() || location.pathname.includes(PLAYER_PATH)) return;
        const task = readTask();
        if (!task) return;
        if (Number(task.resumeAfter) > Date.now()) {
            setStatus('等待平台同步学习进度…');
            return;
        }
        if (task.awaitingPlayer) {
            setStatus('播放窗口学习中，正在等待…');
            return;
        }
        navigationBusy = true;
        try {
            const isListPage = location.hash.includes('/org/course/list');
            if (isListPage) {
                const desiredCategory = task.listCategory === '专业科目' ? '专业科目' : '公需科目';
                const categoryButton = await waitFor(() => findListCategory(desiredCategory), 15000, 500);
                if (!categoryButton) {
                    setStatus(`未找到“${desiredCategory}”分类`, '#ff8c7a');
                    return;
                }
                if (currentListCategory() !== desiredCategory) {
                    setStatus(`正在切换到${desiredCategory}…`);
                    clickLikeUser(categoryButton);
                    await sleep(1500);
                    return;
                }

                const listReady = await waitFor(() => {
                    const cards = Array.from(document.querySelectorAll('.course__list--content > section.course__item, section.course__item'))
                        .filter(isVisible);
                    const pagination = readPagination() || { current: 1, next: null, hasNext: false };
                    return cards.length ? { cards: cards.length, pagination } : null;
                }, 15000, 500);
                if (!listReady) {
                    setStatus(`${desiredCategory}课程卡片尚未加载完成`, '#ffb347');
                    return;
                }
                setStatus(`正在逐项检查${desiredCategory}第 ${listReady.pagination.current} 页，共 ${listReady.cards} 个项目（未购买项目将跳过）…`);
            }

            const course = findFirstIncompleteEnterButton();
            if (course) {
                task.phase = 'must';
                task.listUrl = location.href;
                writeTask(task);
                setStatus(`发现未完成项目：${course.title || '未命名项目'}（${course.progress}%），正在进入…`);
                course.button.scrollIntoView({ block: 'center', behavior: 'smooth' });
                await sleep(350);
                course.button.click();
                return;
            }

            if (isListPage) {
                // 不依赖“进入学习”按钮数量：已完成项目的按钮文案可能会发生变化。
                const pagination = await waitFor(() => readPagination(), 15000, 500);
                if (!pagination) {
                    setStatus('课程列表尚未加载完成，继续等待…', '#ffb347');
                    return;
                }
                if (pagination?.hasNext) {
                    setStatus(`${task.listCategory}第 ${pagination.current} 页已完成，前往下一页…`, '#38d6b3');
                    const previousPage = pagination.current;
                    pagination.next.click();
                    const changed = await waitFor(() => {
                        const latest = readPagination();
                        return latest && latest.current !== previousPage ? latest.current : null;
                    }, 10000, 300);
                    if (!changed) setStatus(`第 ${previousPage + 1} 页切换失败，稍后重试`, '#ff8c7a');
                    else setStatus(`已进入${task.listCategory}第 ${changed} 页`);
                    return;
                }

                if (task.listCategory !== '专业科目') {
                    task.listCategory = '专业科目';
                    task.phase = 'must';
                    writeTask(task);
                    setStatus('公需科目所有页面已完成，切换到专业科目…', '#38d6b3');
                    const professional = findListCategory('专业科目');
                    if (professional) clickLikeUser(professional);
                    await sleep(1500);
                    return;
                }

                setStatus('公需科目和专业科目的所有页面均已完成', '#38d6b3');
                localStorage.removeItem(TASK_KEY);
                return;
            }

            const courseDetail = await waitFor(() =>
                document.querySelector('.text-box, #tab-MUST, #tab-OPTIONAL, .info-content'), 15000
            );
            if (!courseDetail) return;

            // 以页面实时显示的“还需要学习”学分为准，不使用脚本内预设值。
            const creditStatus = await waitForStableCreditStatus();
            if (creditStatus) {
                if (creditStatus.mustRemaining <= 0 && creditStatus.electiveRemaining <= 0) {
                    task.phase = 'must';
                    task.awaitingPlayer = false;
                    writeTask(task);
                    setStatus('页面显示必修、选修剩余学分均为 0，返回总列表', '#38d6b3');
                    await sleep(1000);
                    location.href = task.listUrl || LIST_URL;
                    return;
                }
                task.phase = creditStatus.mustRemaining > 0 ? 'must' : 'elective';
                writeTask(task);
            }

            const isMust = task.phase !== 'elective';
            const phaseName = isMust ? '必修课' : '选修课';
            if (creditStatus) {
                const remaining = isMust ? creditStatus.mustRemaining : creditStatus.electiveRemaining;
                const required = isMust ? creditStatus.mustRequired : creditStatus.electiveRequired;
                setStatus(`${phaseName}还需 ${remaining} 学分${required === null ? '' : `（要求 ${required}）`}，正在选课…`);
            } else {
                setStatus(`未读取到剩余学分，按${phaseName}未完成内容处理…`, '#ffb347');
            }
            const phaseTab = (isMust ? document.getElementById('tab-MUST') : document.getElementById('tab-OPTIONAL')) ||
                findByExactText('.el-tabs__item, [role="tab"], span, div', phaseName);
            if (phaseTab) {
                clickLikeUser(phaseTab);
                await sleep(1000);
            }

            const incompleteCandidates = Array.from(document.querySelectorAll('.btn-item, button, [role="button"], span, div'))
                .filter(element => isVisible(element) && normalize(element.innerText || element.textContent) === '未完成');
            const incompleteFilter = incompleteCandidates.find(element => element.matches('.btn-item, button, [role="button"]')) ||
                incompleteCandidates.find(element => !element.closest('.text-item, .text-info')) ||
                incompleteCandidates[0];
            if (incompleteFilter) {
                clickLikeUser(incompleteFilter);
                await sleep(1500);
            }

            // 详情页内容为异步请求；至少等待课程卡片实际出现后再判断是否为空。
            const cards = await waitFor(() => {
                // Vue 的点击事件绑定在完整的 .text-item.cursor 卡片上，不能取其内部子节点。
                const exactCards = Array.from(document.querySelectorAll('.text-box .text-item.cursor, .text-item.cursor'))
                    .filter(isVisible);
                if (exactCards.length) return exactCards;

                const fallbackCards = Array.from(document.querySelectorAll('.text-box div[class*="text-item"], .info-content div[class*="text-item"]'))
                    .filter(element => isVisible(element) && normalize(element.innerText).length > 0);
                return fallbackCards.length ? fallbackCards : null;
            }, 12000, 600);
            const cardList = cards || [];
            const availableCards = cardList.filter(card => {
                const title = normalize(
                    card.querySelector('.item__name, .section-title, .course-title, .course-name, .text-title, [title], h3, h4')?.getAttribute('title') ||
                    card.querySelector('.item__name, .section-title, .course-title, .course-name, .text-title, [title], h3, h4')?.innerText ||
                    card.innerText
                );
                if (BLOCKED_MICRO_COURSE_TITLES.has(title)) {
                    console.info(`[自动学习] 黑名单课程直接跳过：${title}`);
                    return false;
                }
                return true;
            });
            const target = availableCards.find(card => /未完成|未学习|学习中/.test(normalize(card.innerText))) || availableCards[0];
            if (!target) {
                if (isMust) {
                    task.phase = 'elective';
                    writeTask(task);
                    setStatus('必修课已完成，准备学习选修课', '#38d6b3');
                    window.setTimeout(runNavigation, 1000);
                } else {
                    task.phase = 'must';
                    task.awaitingPlayer = false;
                    writeTask(task);
                    setStatus('本项目必修、选修均已完成，返回总列表', '#38d6b3');
                    await sleep(1000);
                    location.href = task.listUrl || LIST_URL;
                }
                return;
            }

            setStatus(`正在打开${phaseName}播放页…`);
            target.scrollIntoView({ block: 'center', behavior: 'smooth' });
            await sleep(500);
            task.awaitingPlayer = true;
            task.returnUrl = location.href;
            task.resumeAfter = 0;
            writeTask(task);
            // 直接调用课程卡片本身的 click，触发网站绑定在卡片上的 Vue 事件。
            target.click();

            // 若点击没有发生跳转/弹窗，解除等待状态，让下一轮可以重试。
            window.setTimeout(() => {
                const current = readTask();
                if (current?.awaitingPlayer && document.visibilityState === 'visible' && document.hasFocus()) {
                    current.awaitingPlayer = false;
                    writeTask(current);
                    setStatus('播放页未打开，正在重新尝试…', '#ff8c7a');
                    runNavigation();
                }
            }, 8000);
        } finally {
            navigationBusy = false;
        }
    }

    async function waitFor(getter, timeout = 20000, interval = 500) {
        const deadline = Date.now() + timeout;
        while (Date.now() < deadline) {
            try {
                const value = getter();
                if (value && (!('length' in Object(value)) || value.length > 0)) return value;
            } catch (error) {
                // iframe 仍在加载时继续等待。
            }
            await sleep(interval);
        }
        return null;
    }

    function isSectionCompleted(item) {
        return Boolean(item.querySelector('.icon-icon_gouxuan, [class*="gouxuan"], .completed, .is-complete')) ||
            /已完成|已学完/.test(normalize(item.innerText));
    }

    async function playOneSection(frameDoc, item, index) {
        const title = normalize(item.querySelector('.section-title')?.innerText) || `第 ${index + 1} 节`;
        setStatus(`正在播放：${title}`);
        item.scrollIntoView({ block: 'center' });
        item.click();

        let video = await waitFor(() => frameDoc.querySelector('video'), 20000);
        if (!video) {
            console.info(`[自动学习] ${title} 未发现视频，继续下一节。`);
            await sleep(1500);
            return;
        }

        video.muted = true;
        video.volume = 0;
        video.playbackRate = readSpeed();
        try { await video.play(); } catch (error) { console.warn('[自动学习] 自动播放暂时受限：', error); }

        await new Promise(resolve => {
            let finished = false;
            let watchedVideo = video;
            const finish = () => {
                if (finished) return;
                finished = true;
                clearInterval(keepAlive);
                clearInterval(completionCheck);
                clearTimeout(safetyTimeout);
                resolve();
            };
            const keepAlive = setInterval(() => {
                try {
                    // 切换课程后平台可能销毁旧 video，所以每轮都重新寻找当前节点。
                    const videos = Array.from(frameDoc.querySelectorAll('video'));
                    const currentVideo = videos.find(node => {
                        const rect = node.getBoundingClientRect();
                        return node.isConnected && rect.width > 0 && rect.height > 0;
                    }) || videos[videos.length - 1];
                    if (currentVideo) {
                        watchedVideo = currentVideo;
                        watchedVideo.muted = true;
                        watchedVideo.volume = 0;
                        watchedVideo.playbackRate = readSpeed();
                        if (watchedVideo.paused && !watchedVideo.ended) watchedVideo.play().catch(() => {});
                    }
                } catch (error) {}
            }, 800);
            const completionCheck = setInterval(() => {
                if ((watchedVideo && watchedVideo.ended) || isSectionCompleted(item)) finish();
            }, 1000);
            const safetyTimeout = setTimeout(() => {
                console.warn(`[自动学习] ${title} 等待超时，转到下一节。`);
                finish();
            }, 6 * 60 * 60 * 1000);
        });
    }

    async function startAutoPlay() {
        if (autoPlayStarted || !location.pathname.includes(PLAYER_PATH)) return;
        autoPlayStarted = true;
        createPanel();
        setStatus('正在读取课程目录…');

        const frame = await waitFor(() => document.querySelector('#aliPlayerFrame'), 25000);
        if (!frame) {
            setStatus('未检测到播放框架', '#ff8c7a');
            autoPlayStarted = false;
            return;
        }

        let frameDoc;
        try {
            frameDoc = frame.contentWindow.document;
        } catch (error) {
            setStatus('播放器 iframe 跨域，无法读取目录', '#ff8c7a');
            autoPlayStarted = false;
            return;
        }

        const sections = await waitFor(() => Array.from(frameDoc.querySelectorAll('.first-line')), 25000);
        if (!sections || sections.length === 0) {
            setStatus('未解析到课程目录', '#ff8c7a');
            autoPlayStarted = false;
            return;
        }

        setStatus(`发现 ${sections.length} 个小节`);
        for (let index = 0; index < sections.length; index += 1) {
            const item = sections[index];
            if (isSectionCompleted(item)) {
                console.info(`[自动学习] 跳过已完成：${normalize(item.innerText)}`);
                continue;
            }
            await playOneSection(frameDoc, item, index);
            await sleep(1200);
        }

        setStatus('全部小节学习完毕', '#38d6b3');
        await sleep(2000);

        const task = readTask();
        if (task) {
            task.awaitingPlayer = false;
            task.lastPlayerFinishedAt = Date.now();
            task.resumeAfter = Date.now() + 5000;
            writeTask(task);
        }

        const finishedAt = Date.now();
        try {
            flowChannel?.postMessage({ type: 'TM_CQRL_PLAYER_DONE', finishedAt });
        } catch (error) {}
        try {
            if (window.opener && !window.opener.closed) {
                window.opener.postMessage({ type: 'TM_CQRL_PLAYER_DONE', finishedAt }, location.origin);
                if (typeof window.opener.__TM_CQRL_RESUME__ === 'function') {
                    window.opener.__TM_CQRL_RESUME__(finishedAt);
                }
            }
        } catch (error) {
            console.debug('[自动学习] 无法直接通知原页面，将由原页面主动恢复。', error);
        }

        // 新窗口播放时关闭并让原详情页继续；同标签播放时返回之前的详情页。
        const openerAvailable = Boolean(window.opener && !window.opener.closed);
        if (openerAvailable) {
            window.close();
            if (!window.closed) setStatus('已完成，请关闭本页；原页面将自动继续', '#38d6b3');
        } else if (task && task.returnUrl) {
            setStatus('正在返回课程详情页继续学习…', '#38d6b3');
            location.href = task.returnUrl;
        } else {
            setStatus('学习完毕，请返回课程列表', '#38d6b3');
        }
    }

    function init() {
        createPanel();
        if (location.pathname.includes(PLAYER_PATH)) {
            if (taskIsActive()) startAutoPlay();
            else setStatus('等待从课程列表启动全流程');
        }
        else if (taskIsActive()) runNavigation();
    }

    window.addEventListener('load', init);
    window.addEventListener('hashchange', () => window.setTimeout(init, 400));
    window.addEventListener('storage', event => {
        if (event.key === TASK_KEY) window.setTimeout(init, 500);
    });
    window.addEventListener('message', event => {
        if (event.origin === location.origin && event.data?.type === 'TM_CQRL_PLAYER_DONE') {
            resumeAfterPlayer(event.data.finishedAt);
        }
    });
    if (flowChannel) {
        flowChannel.addEventListener('message', event => {
            if (event.data?.type === 'TM_CQRL_PLAYER_DONE') resumeAfterPlayer(event.data.finishedAt);
        });
    }
    // 浏览器可能冻结后台标签；播放窗口关闭、原页面重新获得焦点时立即补一次检查。
    window.addEventListener('focus', () => window.setTimeout(init, 300));
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) window.setTimeout(init, 300);
    });
    window.setInterval(() => {
        createPanel();
        if (location.pathname.includes(PLAYER_PATH)) {
            if (taskIsActive()) startAutoPlay();
        }
        else if (taskIsActive()) runNavigation();
    }, 1200);

    init();
})();
