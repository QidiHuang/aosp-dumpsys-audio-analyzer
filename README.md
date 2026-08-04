# Dumpsys Analyzer

## 1. 简介
出于跨平台兼容性、执行效率以及 UI 体验的考虑，我们将原有的 VBScript 脚本重构成纯 Python (Tkinter) 实现版本。
使用 Python 可以不仅在 Windows，还在 Linux 和 macOS 平台上使用，并且摆脱了对 `wso.dll` 的依赖。
重构后的 `dumpsys_analyzer.py` 合并了 ADB 提取请求，优化了内存字符串正则分析，极大提高了运行速度，降低了内存占用。

## 2. 核心特性
1. **更好的UI显示**：采用现代化的 Tkinter Canvas 绘制清晰的数据拓扑图 (App -> Thread -> Track -> Buffer)；
2. **更快速的日志分析**：一次性抓取 `dumpsys` 和 `ps` 日志，完全在内存中处理正则匹配，消除磁盘 I/O 耗时；
3. **更低的内存占用**：杜绝了高频多次唤起 ADB 子进程造成的内存消耗；
4. **更多的Android版本兼容**：采用弹性正则表达式，动态适配更多 Android 版本输出的差异；
5. **更便于人类阅读的代码风格**：纯 Python 面向对象重构，告别难于维护的传统 VBScript/DLL 调用。

## 3. 依赖环境
- Python 3.6+
- Tkinter (通常随 Python 安装)
- ADB 环境已配置好并在 PATH 中

## 4. 使用说明
**在线分析设备 (实时连接 Android 手机)**：
```sh
python3 dumpsys_analyzer.py
```

**离线分析文本文件**：
```sh
python3 dumpsys_analyzer.py --test-file dumpsys_files/music_play.txt
```
