# 喵程智管

Windows 桌面的**项目进度管理工具**。把项目拆成「阶段」，阶段下再挂「任务」，阶段用流程图展示进展；项目还可以挂文件附件，最后一键导出 Excel 或整包 zip。

单文件 exe，免安装，不用联网，所有数据都存在本地。

## 功能一览

| 功能 | 说明 |
|---|---|
| 三级结构 | 项目 → 阶段 → 任务，任务可写详细描述 |
| 阶段流程图 | 蛇形排版 + 箭头连接，已完成绿色、进行中浅蓝，双击进入阶段 |
| 进度自动计算 | 已完成阶段数 ÷ 总阶段数，项目列表直接显示百分比 |
| 文件附件 | 每个项目可挂任意文件，双击用系统默认程序打开 |
| 搜索 | 项目名实时模糊搜索；按日期搜索（创建日或更新日） |
| 导出 Excel | 一次导出全部数据，含「项目 / 阶段 / 任务 / 文件列表」四张表 |
| 导出完整项目 | 多选项目，打包成一个 zip（Excel + 所有附件） |

## 目录结构

```
喵程智管/
├── reminder_app.py        # 主程序（2024 行，tkinter + sqlite3）
├── dist/
│   └── 喵程智管.exe       # 打包好的程序，11 MB
├── 喵程智管.spec          # PyInstaller 打包配置
├── package.bat            # 一键打包脚本
├── application-note.ico   # 程序图标
├── 操作说明.md            # 完整上手指南（新手看这个）
└── README.md
```

## 下载与运行

打包好的程序在 `dist/喵程智管.exe`，双击即可，无需安装 Python。

**重要：把 exe 放在一个固定文件夹里运行，不要放在压缩包里直接打开、也不要放在桌面来回移动。** 程序会在「当前工作目录」创建数据文件，目录一变，就看不到原来的数据了。

启动后同目录会生成：

```
project_manager.db     ← 数据库：项目 / 阶段 / 任务 / 文件记录
项目文件/              ← 上传的附件副本（按项目名分文件夹）
```

## 数据存放位置（务必理解）

代码里用的是**相对路径**：

```python
sqlite3.connect('project_manager.db')                      # 数据库
os.path.join(os.getcwd(), "项目文件", project_name)         # 附件目录
```

所以数据跟着「启动时的工作目录」走，不是跟着 exe 走。用快捷方式启动时，「起始位置」必须设为 exe 所在目录，否则程序会在别处新建一个空库，看起来像"数据全丢了"。

备份时 `project_manager.db` 和 `项目文件/` **两个都要拷**，只拷数据库的话附件全丢。

## 数据库结构

```sql
projects       (id, name, created_at, updated_at)
stages         (id, project_id, name, start_time, end_time, status)
tasks          (id, stage_id, name, description, created_at)
project_files  (id, project_id, file_name, file_path, file_size, file_type, uploaded_at)
```

- `stages.status` 只有两个值：`进行中`、`已完成`
- `project_files.file_path` 存的是**上传时的绝对路径**，所以换电脑或移动文件后，旧记录打不开
- 阶段和任务都设了 `ON DELETE CASCADE`，删除项目会连带删掉阶段、任务、文件记录（并尝试删除 `项目文件/<项目名>` 整个文件夹）

## 从源码运行

```bash
pip install xlsxwriter
python reminder_app.py
```

除标准库（tkinter、sqlite3）外只有一个第三方依赖：`xlsxwriter`（导出 Excel 用）。

## 重新打包

双击 `package.bat`，或手动执行：

```bash
pyinstaller --onefile --windowed --noconsole --name "喵程智管" --icon=application-note.ico reminder_app.py
```

产物在 `dist/喵程智管.exe`，约 11 MB。打包配置是 `datas=[]`，**不会**把数据库和附件打进 exe，这是有意为之——数据必须留在外面才能备份和迁移。

## 几个必须知道的规则

1. **新增阶段会自动结束上一个阶段**：程序会把「开始时间最晚的那个阶段」状态改成已完成，并写入结束时间。同一时间只能有一个进行中的阶段。
2. **进度只算阶段，不算任务**：一个阶段里 10 个任务做完了，只要阶段没点「结束阶段」，进度就不动。
3. **上传文件按项目名称建目录，一次只能传一个**：重命名项目后再上传，会生成新目录，旧目录留在原处。
4. **同名文件覆盖后列表会出现两条记录**：程序会覆盖磁盘文件，但数据库里新旧两条记录都在，指向同一个路径。删掉其中一条会连物理文件一起删。
5. **导出 Excel 导的是全部项目**，不受搜索框筛选影响。

详细操作见 [操作说明.md](操作说明.md)。

## 许可

个人自用工具，代码可自由参考修改。
