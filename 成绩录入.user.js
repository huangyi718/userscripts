// ==UserScript==
// @name         Excel 上传按钮（支持 iframe 和动态加载）
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  在网页右上角添加 Excel 上传按钮，支持 Vue/React 单页应用和 iframe 动态加载，逐个录入成绩并记录情况
// @author       HUANGYI_CQWU
// @match        https://jwmis.cqwu.edu.cn/*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==
(function () {
    'use strict';

    console.log("📌 脚本已加载，等待 iframe#frmReportA 出现...");

    const observer = new MutationObserver(() => {
        const iframe = document.getElementById("frmReportA");

        if (iframe && !document.getElementById("excelUploadBtn")) {
            console.log("📌 检测到 iframe#frmReportA，插入上传按钮...");

            // **创建上传按钮**
            let fileInput = document.createElement("input");
            fileInput.type = "file";
            fileInput.id = "excelUploadBtn";
            fileInput.accept = ".xlsx, .xls";

            // **设置按钮样式**
            Object.assign(fileInput.style, {
                position: "fixed",
                top: "10px",
                left: "50%",
                transform: "translateX(-50%)",
                zIndex: "9999",
                padding: "10px",
                background: "#007bff",
                color: "#fff",
                border: "none",
                borderRadius: "5px",
                cursor: "pointer",
                fontSize: "16px"
            });

            document.body.appendChild(fileInput);
            console.log("✅ 上传按钮已插入！");

            // **监听 Excel 文件上传**
            fileInput.addEventListener("change", function (event) {
                let file = event.target.files[0];

                // 如果没有文件，跳过
                if (!file) return;

                let reader = new FileReader();

                reader.onload = function (e) {
                    let data = new Uint8Array(e.target.result);
                    let workbook = XLSX.read(data, { type: 'array' });
                    let sheet = workbook.Sheets[workbook.SheetNames[0]];
                    let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    //console.log("📌 Excel 数据解析完成：", jsonData);

                    let students = [];
                    jsonData.forEach((row, index) => {
                        if (row.length >= 3) {
                            let name = row[0].trim();  // 姓名
                            let studentId = String(row[1]).trim(); // 学号
                            let score = row[2]; // 成绩
                            students.push({ name, studentId, score, rowIndex: index + 2 });  // 将行号加到数组中，便于后续处理
                        }
                    });

                    //console.log("📌 解析后的学生数据：", students);

                    // **确保 iframe 加载完成**
                    setTimeout(() => {
                        let iframeDocument = iframe.contentDocument || iframe.contentWindow.document;
                        if (!iframeDocument) {
                            console.error("❌ 无法访问 iframe 的 document（可能是跨域问题）");
                            return;
                        }

                        console.log("📌 开始匹配网页中的学生信息...");

                        // **监听 iframe 内部的 DOM 变化**
                        const iframeObserver = new MutationObserver(() => {
                            processScores(iframeDocument, students);
                        });

                        iframeObserver.observe(iframeDocument, { childList: true, subtree: true });

                        // **首次填充成绩**
                        processScores(iframeDocument, students);
                    }, 20);  // 延时 2 秒，等待 iframe 内容加载
                };

                reader.readAsArrayBuffer(file);

                // **上传完成后，重置文件输入框，确保下一次能重新选择文件**
                fileInput.value = "";
            });

            // **停止 MutationObserver，避免重复执行**
            observer.disconnect();
        }
    });

    // **监听整个页面，等待 iframe 加载**
    observer.observe(document.body, { childList: true, subtree: true });

    /**
     * **逐个录入成绩并记录情况**
     * @param {Document} iframeDocument - iframe 内部的 document 对象
     * @param {Array} students - 解析后的学生数据
     */
    function processScores(iframeDocument, students) {
        let elements = iframeDocument.querySelectorAll("td[name='yhxh']");
        if (elements.length === 0) {
            console.log("⚠️ 未找到学号列，可能是表格未加载完成");
            return;
        }

        //console.log("✅ 成功找到学号列:", elements);

        // **用来保存每一行的记录**
        let resultData = [];

        // **逐个录入成绩**
        elements.forEach(td => {
            let studentId = td.innerText.trim(); // 获取学号
            let rowPrefix = td.id.split("_")[0]; // 获取行的前缀，例如 tr0

            let nameTd = iframeDocument.getElementById(`${rowPrefix}_xm`); // 获取姓名
            if (nameTd) {
                let studentName = nameTd.innerText.trim(); // 获取网页上的姓名
                console.log(`🎯 学号: ${studentId}, 姓名: ${studentName}`);

                // **匹配 Excel 数据**
                let student = students.find(s => s.studentId === studentId && s.name === studentName);
                console.log(`********`);
                console.log(student);
                if (student) {
                    // **找到该学生，录入成绩**
                    let scoreInput = iframeDocument.getElementById(`${rowPrefix}_zhcj_`); // 获取成绩输入框
                    if (scoreInput) {
                        scoreInput.value = student.score;

                        // **手动触发 input 事件**
                        scoreInput.dispatchEvent(new Event("input", { bubbles: true }));

                        console.log(`✅ 成功填充成绩: ${student.score} -> ${rowPrefix}_zhcj_`);

                        // **记录成绩和备注已录入**
                        resultData.push([student.name, student.studentId, student.score, "已录入"]);
                    }
                } else {
                    // **没有该生的成绩**
                    console.log(`❌ 在 Excel 中没有找到学号 ${studentId} 和 姓名 ${studentName} 的成绩！`);

                    // **记录该生没有成绩**
                    resultData.push([studentName, studentId, "", "没有该生的成绩"]);
                }
            }
        });

        // **生成并保存新的 Excel 文件**
        let newWorkbook = XLSX.utils.book_new();
        let newSheet = XLSX.utils.aoa_to_sheet(resultData);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "成绩录入情况");

        // **下载 Excel 文件**
        const newFileName = "成绩录入情况.xlsx";
        XLSX.writeFile(newWorkbook, newFileName);
    }
})();
