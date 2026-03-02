// ==UserScript==
// @name         根据系统的课表对应生成教学日历的日期
// @namespace    http://tampermonkey.net/
// @version      3.0
// @description  完美支持CQWU，教学日历生成
// @author       HY
// @match        *://*.cqwu.edu.cn/*
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/huangyi718/coursedate/course-table-parser.user.js  
// @downloadURL  https://raw.githubusercontent.com/huangyi718/coursedate/course-table-parser.user.js  
// ==/UserScript==

(function() {
    'use strict';
    let observer = null;
    const safeTrim = (str) => typeof str === 'string' ? str.trim() : '';
    let debounceTimer = null;
    let allTdData = [];
    let allParsedCourses = [];
    let exportButton = null;
    // 日历相关变量
    let calendarModal = null;
    let selectedDate = null;
    let currentDisplayDate = new Date();

    // ==================== 1. 基础工具函数 ====================
    // 中文日期格式化
    function formatDateChinese(date) {
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const day = date.getDate();
        return `${year}年${month}月${day}日`;
    }

    // UTC日期转中文（解决时区偏差）
    function formatUTCDateChinese(utcDate) {
        const localDate = new Date(utcDate.getTime() + utcDate.getTimezoneOffset() * 60000);
        return formatDateChinese(localDate);
    }

    // 星期映射（1=周一，7=周日）
    const WEEK_MAP = { 1: '周一', 2: '周二', 3: '周三', 4: '周四', 5: '周五', 6: '周六', 7: '周日' };
    function getWeekDayText(weekDayNum) {
        return WEEK_MAP[weekDayNum] || `未知星期(${weekDayNum})`;
    }

    // 分组排序号提取（无→0，1组→1...）
    function getGroupSortNum(groupStr) {
        if (groupStr === "无" || !groupStr) return 0;
        const numMatch = groupStr.match(/\d+/);
        return numMatch ? parseInt(numMatch[0], 10) : 999;
    }

    // 本地日期格式化（复用）
    function formatLocalDateChinese(date) {
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const day = date.getDate();
        return `${year}年${month}月${day}日`;
    }

    // 本地星期文本（复用）
    function getLocalWeekDayText(weekDayNum) {
        const weekDayMap = {1: '周一', 2: '周二', 3: '周三', 4: '周四', 5: '周五', 6: '周六', 7: '周日'};
        return weekDayMap[weekDayNum] || '未知';
    }


    // ==================== 2. 新增：日历选择器核心逻辑 ====================
    // 1. 创建日历模态框HTML
    function createCalendarModal() {
        if (calendarModal) return;

        // 日历样式（内联避免影响原页面）
        const style = document.createElement('style');
        style.textContent = `
            #courseCalendarModal {
                position: fixed; top: 0; left: 0; right: 0; bottom: 0;
                background: rgba(0,0,0,0.5); display: none; align-items: center; justify-content: center;
                z-index: 99999; margin: 0; padding: 0; box-sizing: border-box;
            }
            #calendarContent {
                background: white; border-radius: 8px; width: 100%; max-width: 400px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.15); overflow: hidden;
            }
            #calendarHeader {
                background: #2196F3; color: white; padding: 16px; text-align: center;
                font-size: 18px; font-weight: 600;
            }
            #calendarNav {
                display: flex; justify-content: space-between; align-items: center;
                padding: 12px 16px; border-bottom: 1px solid #eee;
            }
            #calendarNav button {
                background: transparent; border: none; cursor: pointer;
                width: 36px; height: 36px; border-radius: 50%; display: flex;
                align-items: center; justify-content: center; color: #666;
            }
            #calendarNav button:hover {
                background: #f5f5f5; color: #2196F3;
            }
            #currentMonthText {
                font-size: 16px; font-weight: 500; color: #333;
            }
            .weekTitle {
                display: grid; grid-template-columns: repeat(7, 1fr);
                text-align: center; padding: 8px 0; background: #fafafa;
                border-bottom: 1px solid #eee;
            }
            .weekTitle div {
                font-size: 14px; color: #666; font-weight: 500;
            }
            #calendarGrid {
                display: grid; grid-template-columns: repeat(7, 1fr);
                gap: 4px; padding: 12px;
            }
            .calendarDay {
                width: 100%; height: 40px; display: flex;
                align-items: center; justify-content: center;
                border-radius: 50%; cursor: pointer; font-size: 14px;
                color: #333; position: relative;
            }
            .calendarDay:hover:not(.emptyDay) {
                background: #e3f2fd;
            }
            .calendarDay.selected {
                background: #2196F3; color: white; font-weight: 500;
            }
            .emptyDay {
                cursor: default; color: #eee;
            }
            #calendarFooter {
                display: flex; justify-content: flex-end; gap: 8px;
                padding: 12px 16px; border-top: 1px solid #eee;
            }
            #calendarFooter button {
                padding: 8px 16px; border-radius: 4px; font-size: 14px;
                cursor: pointer; transition: all 0.2s;
            }
            #cancelCalendarBtn {
                border: 1px solid #ddd; background: white; color: #666;
            }
            #cancelCalendarBtn:hover {
                border-color: #ccc; background: #f9f9f9;
            }
            #confirmCalendarBtn {
                border: none; background: #2196F3; color: white;
            }
            #confirmCalendarBtn:hover {
                background: #1976D2;
            }
        `;
        document.head.appendChild(style);

        // 日历HTML结构
        calendarModal = document.createElement('div');
        calendarModal.id = 'courseCalendarModal';
        calendarModal.innerHTML = `
            <div id="calendarContent">
                <div id="calendarHeader">选择当前学期第一周周一日期</div>
                <div id="calendarNav">
                    <button id="prevMonthBtn"><i>←</i></button>
                    <div id="currentMonthText"></div>
                    <button id="nextMonthBtn"><i>→</i></button>
                </div>
                <div class="weekTitle">
                    <div>日</div><div>一</div><div>二</div><div>三</div><div>四</div><div>五</div><div>六</div>
                </div>
                <div id="calendarGrid"></div>
                <div id="calendarFooter">
                    <button id="cancelCalendarBtn">取消</button>
                    <button id="confirmCalendarBtn">确认</button>
                </div>
            </div>
        `;
        document.body.appendChild(calendarModal);

        // 绑定日历事件
        bindCalendarEvents();
    }

    // 2. 绑定日历交互事件
    function bindCalendarEvents() {
        // 上一个月
        document.getElementById('prevMonthBtn').addEventListener('click', () => {
            currentDisplayDate.setMonth(currentDisplayDate.getMonth() - 1);
            generateCalendar();
        });

        // 下一个月
        document.getElementById('nextMonthBtn').addEventListener('click', () => {
            currentDisplayDate.setMonth(currentDisplayDate.getMonth() + 1);
            generateCalendar();
        });

        // 取消选择
        document.getElementById('cancelCalendarBtn').addEventListener('click', () => {
            calendarModal.style.display = 'none';
            selectedDate = null; // 重置选中状态
        });

        // 确认选择
        document.getElementById('confirmCalendarBtn').addEventListener('click', () => {
            if (!selectedDate) {
                alert('请选择一个日期！');
                return;
            }

            // 验证是否为周一
            const selectedDateObj = new Date(selectedDate);
            const localWeekDay = selectedDateObj.getDay() || 7; // 1=周一，7=周日
            if (localWeekDay !== 1) {
                const weekDayMap = {1: '周一', 2: '周二', 3: '周三', 4: '周四', 5: '周五', 6: '周六', 7: '周日'};
                alert(`选择的${selectedDate}是${weekDayMap[localWeekDay]}，请重新选择第一周周一！`);
                return;
            }

            // 执行原有日期计算逻辑
            const coursesWithDate = calculateCourseDates(selectedDate, allParsedCourses);
            if (coursesWithDate.length) exportToExcel(coursesWithDate);

            // 隐藏日历并重置
            calendarModal.style.display = 'none';
            selectedDate = null;
        });
    }

    // 3. 按当前时间设置默认月份（上半年3月，下半年9月）
    function setDefaultCalendarMonth() {
        const now = new Date();
        const currentMonth = now.getMonth(); // 0=1月，11=12月

        // 上半年（1-6月：0-5）默认3月（2），下半年（7-12月：6-11）默认9月（8）
        currentDisplayDate = new Date(now.getFullYear(), currentMonth < 6 ? 2 : 8, 1);
    }

    // 4. 生成日历日期网格
    function generateCalendar() {
        const grid = document.getElementById('calendarGrid');
        const monthTextEl = document.getElementById('currentMonthText');
        grid.innerHTML = '';

        const year = currentDisplayDate.getFullYear();
        const month = currentDisplayDate.getMonth();
        const monthNames = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
        monthTextEl.textContent = `${year}年 ${monthNames[month]}`;

        // 1. 调整星期标题顺序：第一列显示周一，最后一列显示周日
        const weekTitles = document.querySelectorAll('.weekTitle div');
        weekTitles[0].textContent = '一';  // 原周日位置改为周一
        weekTitles[1].textContent = '二';
        weekTitles[2].textContent = '三';
        weekTitles[3].textContent = '四';
        weekTitles[4].textContent = '五';
        weekTitles[5].textContent = '六';
        weekTitles[6].textContent = '日';  // 最后一列显示周日

        // 2. 计算当月第一天在新排列中的位置（以周一为第一天）
        const firstDayOriginal = new Date(year, month, 1).getDay(); // 原始值：0=周日，6=周六
        // 转换为新索引：0=周一，1=周二，...，5=周六，6=周日
        let firstDayIndex;
        if (firstDayOriginal === 0) {
            firstDayIndex = 6; // 原始周日 → 新排列最后一位
        } else {
            firstDayIndex = firstDayOriginal - 1; // 原始周一(1)→新索引0，以此类推
        }

        // 3. 当月总天数
        const totalDays = new Date(year, month + 1, 0).getDate();

        // 4. 添加上月占位日期（根据新索引计算）
        for (let i = 0; i < firstDayIndex; i++) {
            const emptyDay = document.createElement('div');
            emptyDay.className = 'calendarDay emptyDay';
            grid.appendChild(emptyDay);
        }

        // 5. 添加当月日期（按新排列规则）
        for (let day = 1; day <= totalDays; day++) {
            const dateEl = document.createElement('div');
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

            dateEl.className = 'calendarDay';
            dateEl.textContent = day;
            dateEl.dataset.date = dateStr;

            // 标记选中日期
            if (selectedDate === dateStr) {
                dateEl.classList.add('selected');
            }

            // 点击选择日期
            dateEl.addEventListener('click', () => {
                selectedDate = dateStr;
                generateCalendar(); // 重新渲染选中状态
            });

            grid.appendChild(dateEl);
        }
    }

    // 5. 显示日历
    function showCalendar() {
        createCalendarModal(); // 确保日历已创建
        setDefaultCalendarMonth(); // 设置默认月份
        generateCalendar(); // 生成日期
        calendarModal.style.display = 'flex'; // 显示模态框
    }


    // ==================== 3. 核心修复：正则表达式（支持多范围周次+无分组） ====================
    function parseCourses(titleContent, weekDayCol) {
        let courseText = safeTrim(titleContent);
        if (!courseText) {
            console.warn("⚠️ 解析课程：输入内容为空");
            return [];
        }
        // 1. 预处理：统一空格类型+合并连续空格
        courseText = courseText.replace(/ /g, ' ').replace(/\s+/g, ' ').trim();

        const isExperiment = courseText.includes("实验");
        let course = {
            "课程类型": isExperiment ? "实验课" : "理论课",
            "星期列索引": weekDayCol-1,
            "原始内容": titleContent,
            "预处理内容": courseText,
            "课程名称": "",
            "考核类型": "",
            "周次": "",
            "周次列表": [],
            "节次": "",
            "节次前缀": "", // 新增：记录"中午"等前缀
            "节次列表": [],
            "教室": "",
            "校区": "",
            "班级": [],
            "教室类型": "",
            "人数上限": ""
        };

        // 解析课程名称+考核类型
        const nameTypeRegex = /^([^;]+?)\s+(实验|考查|考试)\s+(.*)$/;
        const nameTypeMatch = courseText.match(nameTypeRegex);
        if (!nameTypeMatch) {
            console.error(`❌ 名称+考核类型匹配失败：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["课程名称"] = safeTrim(nameTypeMatch[1]);
        course["考核类型"] = safeTrim(nameTypeMatch[2]);
        let remainingText = safeTrim(nameTypeMatch[3]);

        // 理论课解析（含课程属性）
        if (!isExperiment) {
            const attrRegex = /^(必修课|选修课|任选课)\s+(.*)$/;
            const attrMatch = remainingText.match(attrRegex);
            if (!attrMatch) {
                console.error(`❌ 理论课属性匹配失败：${courseText.substring(0, 40)}...`);
                return [];
            }
            course["课程属性"] = safeTrim(attrMatch[1]);
            remainingText = safeTrim(attrMatch[2]);

        } else {
            course["组数"] = "无"; // 实验课专属字段
        }

        // 解析周次
        const weekRegex = /^\[([\d,\-]+)\]周\s+(.*)$/;
        const weekMatch = remainingText.match(weekRegex);
        if (!weekMatch) {
            console.error(`❌ 周次匹配失败：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["周次"] = safeTrim(weekMatch[1]);
        course["周次列表"] = course["周次"].split(',').flatMap(range => {
            const [start, end] = range.split('-').map(Number);
            return end ? Array.from({ length: end - start + 1 }, (_, i) => start + i) : [start];
        });
        remainingText = safeTrim(weekMatch[2]);

        // 核心修复：支持带"中午"前缀的节次（如"中午2-3节"）
        const sectionRegex = /^(中午|傍晚)?\s*([\d\-]+)节\s+(.*)$/;
        const sectionMatch = remainingText.match(sectionRegex);
        if (!sectionMatch) {
            console.error(`❌ 节次匹配失败（含中午前缀）：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["节次前缀"] = safeTrim(sectionMatch[1] || ""); // 记录"中午"前缀
        course["节次"] = safeTrim(sectionMatch[2]);
        const [secStart, secEnd] = course["节次"].split('-').map(Number);
        course["节次列表"] = Array.from({ length: secEnd - secStart + 1 }, (_, i) => secStart + i);
        remainingText = safeTrim(sectionMatch[3]);

        // 实验课解析组数
        if (isExperiment) {
            const groupRegex = /^(\d+组)\s+(.*)$/;
            const groupMatch = remainingText.match(groupRegex);
            if (!groupMatch) {
                console.error(`❌ 实验课组数匹配失败：${courseText.substring(0, 40)}...`);
                return [];
            }
            course["组数"] = safeTrim(groupMatch[1]);
            remainingText = safeTrim(groupMatch[2]);
        }

        // 解析教室
        const roomRegex = /^(\S+)\s+(.*)$/;
        const roomMatch = remainingText.match(roomRegex);
        if (!roomMatch) {
            console.error(`❌ 教室匹配失败：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["教室"] = safeTrim(roomMatch[1]);
        remainingText = safeTrim(roomMatch[2]);

        // 解析校区
        const campusRegex = /^([^;]+?)\s+(.*)$/;
        const campusMatch = remainingText.match(campusRegex);
        if (!campusMatch) {
            console.error(`❌ 校区匹配失败：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["校区"] = safeTrim(campusMatch[1]);
        remainingText = safeTrim(campusMatch[2]);

        // 解析班级（支持含[]符号的班级名称）
        const classRegex = /^(.+?)\s+([^;]+?)\s+(\d+)$/;
        const classMatch = remainingText.match(classRegex);
        if (!classMatch) {
            console.error(`❌ 班级+类型+人数匹配失败：${courseText.substring(0, 40)}...`);
            return [];
        }
        course["班级"] = classMatch[1].split(';').map(c => safeTrim(c));
        course["教室类型"] = safeTrim(classMatch[2]);
        course["人数上限"] = safeTrim(classMatch[3]);

        return [course];
    }


    // ==================== 4. 日期计算（适配多范围周次） ====================
    function calculateCourseDates(firstMondayInput, courses) {
        // 基础校验：输入为空或无课程数据直接返回
        if (!firstMondayInput || !courses.length) {
            console.log("⚠️ 未输入第一周周一日期或无课程数据，不进行日期计算");
            return [];
        }

        // 1. 日期格式与有效性验证
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dateRegex.test(firstMondayInput)) {
            alert("⚠️ 日期格式错误！请输入「YYYY-MM-DD」（如2024-09-02）");
            return [];
        }

        const [year, month, day] = firstMondayInput.split("-").map(Number);
        const firstMondayDate = new Date(year, month - 1, day);
        if (isNaN(firstMondayDate.getTime())) {
            alert("⚠️ 无效日期！请输入真实存在的日期");
            return [];
        }

        // 2. 课程数据校验（避免无效课程导致计算错误）
        const validCourses = courses.filter(course => {
            const hasValidWeekList = Array.isArray(course.周次列表) && course.周次列表.length > 0;
            const hasValidWeekDay = typeof course.星期列索引 === 'number' && course.星期列索引 >= 1 && course.星期列索引 <= 7;

            if (!hasValidWeekList) {
                console.warn(`⚠️ 课程「${course.课程名称}」周次列表无效，跳过计算`);
            }
            if (!hasValidWeekDay) {
                console.warn(`⚠️ 课程「${course.课程名称}」星期索引无效（必须1-7），跳过计算`);
            }
            return hasValidWeekList && hasValidWeekDay;
        });

        if (validCourses.length === 0) {
            console.log("⚠️ 无有效课程数据，不进行日期计算");
            return [];
        }

        // 3. 计算每节课的具体日期（使用本地时间）
        const resultCourses = [];
        validCourses.forEach(course => {
            course.周次列表.forEach(weekNum => {
                const daysOffset = (weekNum - 1) * 7;
                const currentWeekMonday = new Date(firstMondayDate);
                currentWeekMonday.setDate(firstMondayDate.getDate() + daysOffset);

                const weekDayOffset = course.星期列索引 - 1;
                const courseDate = new Date(currentWeekMonday);
                courseDate.setDate(currentWeekMonday.getDate() + weekDayOffset);

                const dateChinese = formatLocalDateChinese(courseDate);
                const weekDayText = getLocalWeekDayText(course.星期列索引);
                const dateWithWeek = `${dateChinese}（${weekDayText}）`;
                const rawDate = courseDate.toISOString().split('T')[0];

                const sameCourseCount = resultCourses.filter(c =>
                    c.课程名称 === course.课程名称 &&
                    c.班级.join(';') === course.班级.join(';') &&
                    c.组数 === course.组数
                ).length + 1;

                resultCourses.push({
                    ...course,
                    周次序号: weekNum,
                    周次描述: `第${weekNum}周`,
                    原始日期: rawDate,
                    中文日期: dateChinese,
                    带星期日期: dateWithWeek,
                    总课时: course.周次列表.length,
                    进度: `${sameCourseCount}/${course.周次列表.length}`
                });
            });
        });

        //console.log(`✅ 日期计算完成，共生成 ${resultCourses.length} 条课程日期数据`);
        return resultCourses;
    }


    // ==================== 5. 二列排序（保持逻辑） ====================
    function sortByFourColumns(courses) {
        const validCourses = courses.filter(course =>
            course && typeof course === 'object'
        );
    
        // 新逻辑：先按课程名称→再按班级→最后按原始日期排序
        return validCourses.sort((a, b) => {
            // 1. 先比较课程名称（中文拼音排序）
            const nameA = a.课程名称 || '';
            const nameB = b.课程名称 || '';
            const nameCompare = nameA.localeCompare(nameB, 'zh-CN');
            if (nameCompare !== 0) {
                return nameCompare;
            }
    
            // 2. 课程名称相同时，比较班级（将班级数组转为字符串后比较）
            // 班级可能是数组（如["23计科1", "23计科2"]），用分号拼接为字符串再比较
            const classA = Array.isArray(a.班级) ? a.班级.join(';') : (a.班级 || '');
            const classB = Array.isArray(b.班级) ? b.班级.join(';') : (b.班级 || '');
            const classCompare = classA.localeCompare(classB, 'zh-CN');
            if (classCompare !== 0) {
                return classCompare;
            }
    
            // 3. 班级也相同时，最后比较原始日期（YYYY-MM-DD格式）
            const dateA = a.原始日期 || '';
            const dateB = b.原始日期 || '';
            return dateA.localeCompare(dateB);
        });
}


    // ==================== 6. Excel导出（包含所有字段） ====================
    function exportToExcel(courses) {
        if (!courses.length) {
            alert("❌ 无课程数据可导出！");
            return;
        }
        // 在排序前输出原始课程数据
        //console.log("排序前的课程数据：", courses);
        // 1. 排序
        const sortedCourses = sortByFourColumns(courses);
        // 2. 去重
        const deduplicatedCourses = removeDuplicateCourses(sortedCourses);

        // 3. 表头与数据处理
        const headers = [
            "课程名称", "课程类型", "课程属性", "原始周次",
            "总课时", "当前周次", "班级", "组数",
            "课程日期", "星期", "节次", "上课教室",
            "上课校区", "教室类型", "课程人数"
        ];

        const rows = [headers.join(',')];
        deduplicatedCourses.forEach(course => {
            const getSafeValue = (key, isArray = false) => {
                const value = course[key] ?? "";
                if (isArray) return Array.isArray(value) ? value.join(';') : String(value);
                 // 对“节次和周次”字段特殊处理：前加英文单引号，强制Excel按文本显示
                if (key === "节次"||key === "周次") {
                    return `'${String(value)}`.replace(/"/g, '""'); // 单引号+转义双引号
                }
                return String(value).replace(/"/g, '""');
            };

            const row = [
                `"${getSafeValue("课程名称")}"`,
                `"${getSafeValue("课程类型")}"`,
                `"${getSafeValue("课程属性")}"`,
                `"${getSafeValue("周次")}"`,
                `"${getSafeValue("总课时")}"`,
                `"${getSafeValue("周次序号")}"`,
                `"${getSafeValue("班级", true)}"`,
                `"${getSafeValue("组数")}"`,
                `"${getSafeValue("中文日期")}"`,
                `"${getWeekDayText(getSafeValue("星期列索引"))}"`,
                `"${getSafeValue("节次")}"`,
                `"${getSafeValue("教室")}"`,
                `"${getSafeValue("校区")}"`,
                `"${getSafeValue("教室类型")}"`,
                `"${getSafeValue("人数上限")}"`
            ].join(',');

            rows.push(row);
        });

        // 4. 生成CSV并下载
        const csvContent = "\uFEFF" + rows.join('\n');
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');

        const firstCourseClass = deduplicatedCourses[0]?.班级;
        const className = Array.isArray(firstCourseClass)
            ? firstCourseClass[0]
            : firstCourseClass || "未知班级";
        const today = formatLocalDateChinese(new Date());
        a.download = `课程表_${className}_${today}.csv`;

        a.href = url;
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 100);

        // 5. 导出统计
        //alert(`✅ 导出成功！\n- 排序后总记录数：${sortedCourses.length}条\n- 去重后总记录数：${deduplicatedCourses.length}条\n- 删除重复记录数：${sortedCourses.length - deduplicatedCourses.length}条\n- 文件名：课程表_${className}_含多周次_去重后_${today}.csv`);
    }

    // 去重函数（用户原有逻辑，暂不启用）
    function removeDuplicateCourses(sortedCourses) {
        const uniqueMap = new Map();
        sortedCourses.forEach((course, index) => {
            const courseName = String(course["课程名称"] ?? "");
            const group = String(course["组数"] ?? "");
            const date = String(course["原始日期"] ?? "");
            const uniqueKey = `${courseName}||${group}||${date}`;
            if (!uniqueMap.has(uniqueKey)) {
                uniqueMap.set(uniqueKey, course);
            }
        });
        return Array.from(uniqueMap.values());
    }


    // ==================== 7. 表格数据提取（含课程分割） ====================
    function extractAndStoreTdData(td, rowIdx, correctedColIdx, tableIdx) {
        // 优先提取font[title]隐藏内容
        let fullTdContent = '';
        const fontElement = td.querySelector('font[title]');
        if (fontElement && fontElement.title) {
            fullTdContent = safeTrim(fontElement.title.replace(/[\n\r]/g, ' '));
        } else {
            fullTdContent = safeTrim(td.textContent);
        }

        if (!fullTdContent) return;

        // 课程分割逻辑
        const processedContent = fullTdContent
            .replace(/\s+/g, ' ')
            .replace(/(\d)([\u4e00-\u9fa5]+?)\s+(考查|考试|实验)/g, '$1 $2 $3');

        const splitMarker = '###SPLIT###';
        const markedContent = processedContent.replace(
            /(\d+)\s+(?=[^0-9]+?\s+(考查|考试|实验))/g,
            (_, num) => `${num}${splitMarker}`
        );

        const splitCourses = markedContent.split(splitMarker)
            .map(course => course.trim())
            .filter(course => course && course.includes('周') && course.includes('节'));

        // 存储TD数据并解析课程
        const currentTdData = {
            "表格索引": tableIdx + 1,
            "行号": rowIdx,
            "修正列号": correctedColIdx,
            "TD完整内容": fullTdContent,
            "分割课程数量": splitCourses.length,
            "分割课程列表": splitCourses
        };
        allTdData.push(currentTdData);

        // 解析分割后的课程
        splitCourses.forEach(course => {
            const parsedCourses = parseCourses(course, correctedColIdx);
            if (parsedCourses.length) {
                allParsedCourses.push(...parsedCourses);
            }
        });
    }

    // 处理表格合并单元格
    function processTable(table, tableIdx) {
        const rows = Array.from(table.querySelectorAll('tr'));
        if (!rows.length) return;

        let mergeRowRecords = [];
        rows.forEach((row, rowIdx) => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            let currentColIdx = 0;

            // 处理已存在的合并行
            for (let i = 0; i < mergeRowRecords.length; i++) {
                const merge = mergeRowRecords[i];
                if (merge.row === rowIdx) {
                    currentColIdx++;
                    merge.remain--;
                    if (merge.remain === 0) mergeRowRecords.splice(i, 1);
                }
            }

            // 遍历单元格
            cellLoop:
            for (const cell of cells) {
                // 跳过被合并占用的列
                for (const merge of mergeRowRecords) {
                    if (merge.row === rowIdx && currentColIdx === merge.col) {
                        currentColIdx++;
                        continue cellLoop;
                    }
                }

                const colSpan = cell.colSpan || 1;
                const rowSpan = cell.rowSpan || 1;
                const correctedColIdx = currentColIdx;

                // 过滤表头和时间槽
                const cellText = safeTrim(cell.textContent);
                const isHeader = /星期[一二三四五六日]/.test(cellText);
                const isTimeSlot = cellText.includes('上午') || cellText.includes('下午');
                if (!isHeader && !isTimeSlot) {
                    extractAndStoreTdData(cell, rowIdx, correctedColIdx, tableIdx);
                }

                // 记录新的合并行
                if (rowSpan > 1) {
                    for (let i = 1; i < rowSpan; i++) {
                        mergeRowRecords.push({
                            row: rowIdx + i,
                            col: currentColIdx,
                            remain: rowSpan - i
                        });
                    }
                    mergeRowRecords.sort((a, b) => a.row - b.row || a.col - b.col);
                }

                currentColIdx += colSpan;
            }
        });
    }


    // ==================== 8. 导出按钮创建（替换为日历触发） ====================
    function createExportButton() {
        if (exportButton) exportButton.remove();
        if (allParsedCourses.length === 0) return;

        exportButton = document.createElement('button');
        exportButton.innerText = `📊 导出课程（${allParsedCourses.length}条）`;
        exportButton.style.cssText = `
            position: fixed; bottom: 30px; right: 30px;
            padding: 12px 20px; background: #2196F3; color: white;
            border: none; border-radius: 8px; font-size: 14px; cursor: pointer;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2); z-index: 9999;
            opacity: 0; transform: translateY(20px); transition: all 0.3s;
        `;

        // 点击按钮显示日历（替换原prompt）
        exportButton.addEventListener('click', showCalendar);

        // hover效果
        exportButton.addEventListener('mouseover', () => {
            exportButton.style.background = "#1976D2";
            exportButton.style.transform = "translateY(0) scale(1.05)";
        });
        exportButton.addEventListener('mouseout', () => {
            exportButton.style.background = "#2196F3";
            exportButton.style.transform = "translateY(0) scale(1)";
        });

        document.body.appendChild(exportButton);
        setTimeout(() => {
            exportButton.style.opacity = 1;
            exportButton.style.transform = "translateY(0)";
        }, 100);
    }

    function outputParsedCourses() {
        if (!allParsedCourses.length) {
            console.log("❌ 未解析到课程数据");
            return;
        }
        createExportButton();
    }


    // ==================== 9. iframe处理 ====================
    function handleFrmReport(iframe) {
        // 监听iframe重新加载
        let frameSrc = iframe.src;
        const srcObserver = new MutationObserver(() => {
            if (iframe.src !== frameSrc) {
                console.log("🔄 iframe重新加载，重新解析");
                srcObserver.disconnect();
                handleFrmReport(iframe);
            }
        });
        srcObserver.observe(iframe, { attributes: true, attributeFilter: ['src'] });

        // 加载完成处理
        iframe.onload = function() {
            try {
                allTdData = [];
                allParsedCourses = [];

                const innerDoc = iframe.contentDocument || iframe.contentWindow.document;
                let tables = innerDoc.querySelectorAll('#mytable0');
                if (!tables.length) tables = innerDoc.querySelectorAll('table');

                if (!tables.length) return;

                tables.forEach((table, idx) => processTable(table, idx));
                outputParsedCourses();

            } catch (e) {
                console.error("❌ 处理iframe失败：", e.message);
            }
        };

        // 已加载则直接处理
        try {
            const innerDoc = iframe.contentDocument;
            if (innerDoc && innerDoc.readyState === "complete") iframe.onload();
        } catch (e) {
            console.warn("⚠️ 跨域限制：无法访问iframe内容");
        }
    }

    function scanForFrmReport(parentDoc) {
        parentDoc = parentDoc || document;
        const iframes = parentDoc.querySelectorAll('iframe');

        iframes.forEach(iframe => {
            if (iframe.id === "frmReport") {
                handleFrmReport(iframe);
                return;
            }

            try {
                const childDoc = iframe.contentDocument;
                if (childDoc && childDoc.readyState === "complete") {
                    scanForFrmReport(childDoc);
                }
            } catch (e) {
                console.log(`⚠️ 跳过跨域iframe：${e.message.slice(0, 30)}...`);
            }
        });
    }


    // ==================== 10. 启动与监听 ====================
    function isTargetRelatedChange(mutations) {
        const ignoreTags = ['INPUT', 'TEXTAREA', 'SELECT', 'SPAN', 'DIV:not([class="div1"])', 'FONT'];
        for (const mutation of mutations) {
            if (mutation.type !== 'childList') continue;
            const hasTargetNode = Array.from(mutation.addedNodes).some(node => {
                return node.tagName === 'IFRAME' || node.tagName === 'TABLE';
            });
            const hasIgnoreNode = Array.from(mutation.addedNodes).some(node => {
                if (!(node instanceof HTMLElement)) return false;
                return ignoreTags.some(tag => node.tagName === tag.toUpperCase() || node.closest(tag));
            });
            if (hasTargetNode && !hasIgnoreNode) return true;
        }
        return false;
    }

    function debounceScan(mutations) {
        if (!isTargetRelatedChange(mutations)) return;
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            if (exportButton) exportButton.remove();
            scanForFrmReport();
        }, 150);
    }

    function start() {
        scanForFrmReport();
        observer = new MutationObserver(debounceScan);
        observer.observe(document.body, { childList: true, subtree: true });

        // 1小时后停止监听
        setTimeout(() => {
            observer.disconnect();
            console.log("\n⏰ 1小时超时，停止监听");
        }, 3600000);
    }

    // 启动脚本
    start();
})();
