# dumpsys_analyzer readme

---

Created by _Qidi_ at _2021/10/09_

---

1. 简介

    出于产品轻量化的考虑，最后我们会发布一个脚本。它主要由 _VBS (VisualBasicScript) / JS (JavaScript)_ 写成，用于周期性地采集和分析 `dumpsys` 信息。再利用 _WSO (WindowsSystemObject)_ 库，将分析后的信息以图形化的形式显示出来。
    这样做的好处是：开发和使用都无需安装任何运行环境；不用编译；易于修改；

2. 目录说明

	+ **developer_manuals**	- 开发者手册，包含 `wso.dll` 库的 API 说明和 VBS 语法说明；
	+ **dumpsys_files**		- 从实机上 dump 到的文本；
	+ **prebuilt_libs**		- `wso.dll` 库文件及注册方法；
    + **src**				- 项目源文件目录
    + **tests**				- 功能测试源文件目录
    + **AudioFlinger_dumpsys_pre_analysis.txt** - AudioFlinger dumpsys 信息预分析结果；
    + **UI_View_Draft.jpg**	- 初步 UI 设计图；


3. 功能子模块说明

	+ `wso.dll` 注册状态检测模块；
	+ `dumpsys` 信息采集和分析模块；
	+ UI 绘制和更新模块；

4. 使用说明

	+ 克隆完整工程；
	+ 通过 `src/main.bat` 启动。
