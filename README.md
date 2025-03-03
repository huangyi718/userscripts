本脚本仅作为浏览器插件使用，不涉及任何违反互联网规定的内容，也无需科学上网。

## 安装步骤

### 如果已安装篡改猴（Tampermonkey）
直接跳转至 [第 3 步](#3-安装脚本)。

### 1. 安装篡改猴（Tampermonkey）扩展

1). 打开浏览器，访问以下网址：
   
   [篡改猴扩展下载](https://microsoftedge.microsoft.com/addons/detail/%E7%AF%A1%E6%94%B9%E7%8C%B4/iikmkjmpaadaobahmlepeloendndfphd)
   
2). 将该扩展添加到浏览器。

### 2. 启用开发者模式

1). 点击浏览器右上角的 **三点菜单**，选择 **扩展**。
   
   ![image](https://github.com/huangyi718/userscripts/blob/main/step/1.png)
   
2). 进入 **管理扩展**。
   
   ![image](https://github.com/huangyi718/userscripts/blob/main/step/2.png)
   
3). **开启开发者模式**。
   
   ![image](https://github.com/huangyi718/userscripts/blob/main/step/3.png)

### 3. 安装脚本

在浏览器地址栏输入以下地址并访问：

[点击安装脚本](https://cdn.jsdelivr.net/gh/huangyi718/userscripts@main/%E6%88%90%E7%BB%A9%E5%BD%95%E5%85%A5.user.js)

### 4. 使用脚本

安装成功后，在成绩系统中登记成绩时，会看到新增的 **上传成绩** 选项卡。

### 5. 上传成绩说明

- 上传的 **Excel 文件格式**：
  - **第一列**：姓名
  - **第二列**：学号
  - **第三列**：成绩

- 脚本会自动比对 **成绩系统中的姓名和学号**，匹配成功后录入成绩。
- 若姓名或学号不匹配，系统会输出日志并保存至 Excel 文件，建议在完成录入后 **下载 Excel 进行核对**。
- 可选择 **暂存成绩**，在最终录入前与原始成绩表进行对比。

---

如果在安装或使用过程中遇到问题，欢迎提交 Issue 反馈！
