// ==UserScript==
// @name         Excel 成绩填充（姓名+学号匹配 + 状态写回）
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  读取 Excel，匹配姓名和学号后填充成绩，并直接修改原 Excel
// @author       You
// @match        https://jwmis.cqwu.edu.cn/*
// @grant        none
// @run-at       document-end
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function() {
    'use strict';

    console.log("📜 脚本已加载，等待 iframe#frmReportA 出现...");

    // 监听 DOM 变化，等待 iframe 出现
    const observer = new MutationObserver(() => {
        const iframe = document.getElementById("frmReportA");

        if (iframe && !document.getElementById("excelUploadBtn")) {
            console.log("✅ iframe#frmReportA 已加载，插入上传按钮...");

            // 1️⃣ 创建上传按钮
            let fileInput = document.createElement("input");
            fileInput.type = "file";
            fileInput.id = "excelUploadBtn";
            fileInput.accept = ".xlsx, .xls";

            // 设置按钮样式（顶部居中）
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

            console.log("📂 上传按钮已成功插入！");

            // 2️⃣ 监听文件上传
            fileInput.addEventListener("change", function(event) {
                let file = event.target.files[0];
                let reader = new FileReader();

                reader.onload = function(e) {
                    let data = new Uint8Array(e.target.result);
                    let workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                    let sheet = workbook.Sheets[workbook.SheetNames[0]];
                    let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    console.log("📊 读取 Excel 数据：", jsonData);

                    // 解析数据（假设 Excel 第一列是姓名，第二列是学号，第三列是成绩）
                    let students = {};
                    jsonData.forEach((row, index) => {
                        if (row.length >= 3) {
                            let name = row[0].trim();  // 姓名
                            let studentId = String(row[1]).trim(); // 学号（转字符串）
                            let score = row[2]; // 成绩
                            students[studentId] = { name, score, index };
                        }
                    });

                    console.log("📋 解析后的学生数据：", students);

                    // 遍历网页，匹配学号和姓名，填充成绩
                    let updateExcel = Array(jsonData.length).fill("未匹配"); // 初始化状态列

                    document.querySelectorAll("td[name='yhxh']").forEach(td => {
                        let studentId = td.innerText.trim(); // 获取学号
                        let rowPrefix = td.id.replace("_yhxh", ""); // 获取行前缀（如 tr0）
                        let nameTd = document.getElementById(`${rowPrefix}_xm`); // 获取姓名单元格

                        if (nameTd) {
                            let studentName = nameTd.innerText.trim(); // 获取网页中的姓名

                            if (students[studentId] && students[studentId].name === studentName) {
                                let scoreInput = document.getElementById(`${rowPrefix}_cj`); // 获取成绩输入框

                                if (scoreInput) {
                                    scoreInput.value = students[studentId].score;
                                    console.log(`✅ 已填充学号 ${studentId}（${studentName}）的成绩: ${students[studentId].score}`);

                                    // Excel 写入 "已完成"
                                    updateExcel[students[studentId].index] = "已完成";
                                } else {
                                    console.warn(`❌ 未找到学号 ${studentId} 的成绩输入框`);
                                    updateExcel[students[studentId].index] = "未找到成绩输入框";
                                }
                            } else {
                                console.warn(`⚠️ 学号 ${studentId} 姓名不匹配，Excel: ${students[studentId]?.name}，网页: ${studentName}`);
                                updateExcel[students[studentId].index] = "信息不匹配";
                            }
                        }
                    });

                    // 3️⃣ 把 "状态" 列写回 Excel 并填充未匹配的单元格红色
                    jsonData.forEach((row, index) => {
                        row[3] = updateExcel[index]; // 状态列

                        // 如果未匹配，则给 Excel 该单元格填充红色
                        if (row[3] !== "已完成") {
                            let cellRef = XLSX.utils.encode_cell({ r: index, c: 3 }); // 状态列（第 4 列）
                            if (!sheet[cellRef]) sheet[cellRef] = {};
                            sheet[cellRef].s = { fill: { fgColor: { rgb: "FF0000" } } }; // 填充红色
                        }
                    });

                    // 4️⃣ 直接覆盖原 Excel 文件
                    XLSX.writeFile(workbook, file.name); // 直接修改原文件
                    console.log("📂 Excel 文件已更新并覆盖！");
                };

                reader.readAsArrayBuffer(file);
            });

            // 停止监听，避免重复执行
            observer.disconnect();
        }
    });

    // 监听整个页面
    observer.observe(document.body, { childList: true, subtree: true });

})();
