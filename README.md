# AutoTask-UI-  

基于 PySide6 的零代码自动化桌面操作工具原型。  
将日常「固定步骤」抽象为可配置的「任务」，由 PyAutoGUI 在后台精确复现，支持定时调度、多任务管理与热插拔式启停。

------

目录  

1. 功能特性  
2. 快速开始  
3. 界面速览  
4. 任务配置示例  
5. 技术栈  
6. 路线图  
7. 贡献指南  

------

1. 功能特性

------

- **可视化任务编排**  
  拖拽式步骤表格，支持「点击 / 双击 / 右键 / 滚轮 / 键盘输入 / 等待」等原子操作。
- **多任务管理**  
  新建、启动、停止、重命名、复制、删除、一键批量启停。
- **定时调度**  
  单次、每日、循环间隔三种策略；
- **图片定位**  
  基于 PyAutoGUI 的模板匹配，自动寻找按钮或区域，无惧分辨率变化。
- **热插拔配置文件**  
  JSON/YAML 双格式，文件即任务，方便版本控制与团队协作。

------

2. 快速开始

------

```bash
# 1. 克隆仓库
git clone https://github.com/yourname/AutoTask-UI.git
cd AutoTask-UI

# 2. 创建虚拟环境（可选）
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt

```

首次启动即出现主界面，可立即创建演示任务。

------

3. 界面速览

------

```
┌────────────────────────────────────────────────────────────┐
│ 任务列表          |  任务详情（标签页切换）                       │
│ ---------------   |  ----------------------------------------- │
│ ▢ 任务 1          |  基本信息  步骤配置  定时策略                │
│ ▢ 任务 2          |  ┌-------------------------------┐         │
│ ▢ 任务 3          |  │ 任务名称: [ Demo            ] │         │
│                   |  │ 描述:     [ 文本框...       ] │         │
│ [+ 新建]          |  │ 步骤表格:                    │         │
│ [▶ 启动] [■ 停止] |  │ 类型| 参数 | 备注            │         │
│ [📋 副本] [🗑 删除] |  └-------------------------------┘         │
└────────────────────────────────────────────────────────────┘
```

- **左侧列表**：右键支持更多操作。  
- **步骤表格**：双击单元格编辑；工具栏可导入截图自动填入路径。  
- **定时策略**：打开「启用定时」后，可设置「每日 09:00」或「每 30 分钟」。

------

4. 技术栈

------

- **GUI**：PySide6（Qt6 官方 Python 绑定）  
- **自动化**：PyAutoGUI + OpenCV（模板匹配）  
- **配置**：JSON 序列化（PyYAML）  
- **调度**：APScheduler（待集成）  
- **打包**：PyInstaller（提供一键 `AutoTask.exe`）

------

## License

MIT © 2024 AutoTask-UI Contributors
------

## bug解决：

* File "I:\Users\pc\anaconda3\envs\AutoTask-UI-\lib\site-packages\pyautogui\__init__.py", line 174, in wrapper     raise ImageNotFoundException  # Raise PyAutoGUI's ImageNotFoundException. pyautogui.ImageNotFoundException

[https://blog.csdn.net/m0_53911267/article/details/134731286]()