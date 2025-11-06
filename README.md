# 工时记录系统

这是一个基于Flask的工时记录和统计平台，支持用户记录和管理每日工作时间。

## 功能特点

- 用户认证系统
- 实时计时器功能
- 每日工时记录
- 周视图展示
- 工时手动补登
- 用户管理（管理员）
- 响应式界面设计

## 技术栈

- 后端：Flask 2.3.3
- 数据库：SQLite + SQLAlchemy ORM
- 认证：Flask-Login
- 前端：Bootstrap 5.1.3
- 数据处理：Pandas + Openpyxl

## 安装说明

1. 克隆仓库：

```bash
git clone https://github.com/yourusername/workingTime.git
cd workingTime
```

2. 创建虚拟环境（推荐）：

```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate  # Windows
```

3. 安装依赖：

```bash
pip install -r requirements.txt
```

4. 运行应用：

```bash
python app.py
```

5. 访问应用：

打开浏览器访问 http://127.0.0.1:5000

## 首次使用

1. 首次访问时会自动进入安装页面
2. 点击"开始安装"按钮完成初始化
3. 使用默认管理员账户登录：
   - 用户名：admin
   - 密码：admin
4. 登录后请立即修改默认密码

## 使用说明

### 仪表盘

- 实时计时器：选择项目并开始计时
- 今日记录：查看当天已记录的工时

### 周视图

- 查看一周工时分布
- 点击单元格添加或编辑工时记录
- 支持上下周切换

### 用户管理（仅管理员）

- 添加新用户
- 删除现有用户
- 查看用户列表

## 注意事项

- 工时记录精确到0.5小时
- 单次记录最短0.5小时，最长8小时
- 系统自动区分上午/下午时段
- 数据库文件存储在instance目录

## 开发说明

- instance/目录用于存放数据库文件
- static/目录用于存放静态资源
- templates/目录包含所有HTML模板
- app.py为主应用文件

## 贡献指南

1. Fork本仓库
2. 创建特性分支
3. 提交更改
4. 推送到分支
5. 创建Pull Request

## 许可证

[MIT License](LICENSE)