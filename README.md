🎉 AutoTask-UI- ✨  
基于 PySide6 的零代码自动化桌面操作工具原型 🚀  
把每天都要重复的「固定步骤」抽象成可拖拽的「任务」🧩，让 PyAutoGUI 在后台帮你精确复现 🎯，还能定时调度 ⏰，一键批量管理 📦！

---

📑 目录  
1. 🌟 功能特性  
2. 🚀 快速开始  
3. 🖥 界面速览  
4. 🧪 任务配置示例  
5. 🛠 技术栈  
6. 🗺 路线图  
7. 🤝 贡献指南  

---

## 1. 🌟 功能特性

| 特性 | Emoji | 说明 |
|---|---|---|
| 可视化任务编排 | 🖱️🎨 | 拖拽式步骤表格，支持「点击 / 双击 / 右键 / 滚轮 / 键盘输入 / 等待」等原子操作 |
| 多任务管理 | 📋🔁 | 新建 ➕、启动 ▶️、停止 ⏹️、重命名 ✏️、复制 📑、删除 🗑️ |
| 定时调度 | ⏰📅 | 单次 / 每日 / 循环间隔 三种策略 |
| 图片定位 | 🔍🖼️ | 基于 PyAutoGUI + OpenCV 模板匹配，无惧分辨率变化 |
| 热插拔配置文件 | 📁🔗 | JSON 格式，文件即任务，Git 友好 |

---

## 2. 🚀 快速开始

```bash
# 1️⃣ 克隆仓库
git clone https://github.com/junior6666/AutoTask-UI-.git
cd AutoTask-UI

# 2️⃣ 创建虚拟环境（可选）
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# 3️⃣ 安装依赖
pip install -r requirements.txt
```

首次启动即可看到主界面，立即创建演示任务！🥳

---

## 3. 🖥 界面速览

![初始化UI.png](img/%E5%88%9D%E5%A7%8B%E5%8C%96UI.png)
- 左侧列表：右键解锁更多姿势 🖱️  
- 步骤表格：双击单元格编辑 ✏️，工具栏导入截图自动填路径 📸  
- 定时策略：启用后，可设置「每日 09:00」或「每 30 分钟」⏲️ 等等

---

## 4. 🧪 任务配置示例

| 步骤类型 | 参数示例 | 备注 |
|---|---|---|
| 🔘 点击 | `btn_ok.png` | 找到按钮并单击 |
| ⌨️ 输入 | `Hello, AutoTask!` | 在当前焦点处输入 |
| ⏱️ 等待 | `2.5` | 等待 2.5 秒 |
| 🖱️ 滚轮 | `-3` | 向下滚动 3 格 |

---

## 5. 🛠 技术栈

- GUI：🐍 PySide6  
- 自动化：🤖 PyAutoGUI + 🔍 OpenCV  
- 配置：📄 JSON / YAML (PyYAML)  
- 调度：⏲️ APScheduler（待集成）  
- 打包：📦 PyInstaller → 一键 `AutoTask.exe`

---

## 6. 🗺 路线图

- ✅ 基础任务编排
- 📊 运行日志 & 可视化报告  
- 🌐 Web 远程控制  (TODO)
- 🧩 插件市场 （TODO）

---

## 7. 🤝 贡献指南

1. Fork ➕ ⭐  
2. 新建分支 `feat/awesome-feature`  
3. 提交 PR 🚀  
4. 等待 Review & Merge 🎊

---

## 📜 License  
MIT © 2024 AutoTask-UI Contributors 🧑‍💻👩‍💻

---

## 📦 项目打包

```bash
pyinstaller -F -w -i icon.ico --add-data "img;img;" main_plus.py

pyinstaller main_plus.spec

pyinstaller -F -w -i icon.ico --add-data "img;img" --name auto_Task2.0.2 main_plus.py
```

---

## 🐞 Bug 解决

**错误**：`pyautogui.ImageNotFoundException`  
**原因**：PyAutoGUI ≥ 0.1.30 在找不到图像时会直接抛异常  
**解决**：降级到 0.1.29

```bash
pip uninstall pyautogui
pip install pyautogui==0.1.29
```

🔗 参考：[CSDN 解决方案](https://blog.csdn.net/m0_53911267/article/details/134731286)