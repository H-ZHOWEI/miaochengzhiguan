import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import sqlite3
from datetime import datetime
from collections import deque
import xlsxwriter  # 新增导入xlsxwriter库用于Excel导出
import os  # 新增导入os模块用于路径处理
import shutil  # 新增导入shutil用于文件操作
import subprocess  # 新增导入subprocess用于打开文件
import zipfile  # 新增导入zipfile用于创建压缩包
import sys  # 新增导入sys模块


class ProjectManager:
    def __init__(self, root):
        self.root = root
        self.root.title("喵程智管 - 版权所有 © 爱吃蟹黄先生")
        self.root.geometry("1200x800")

        # 设置应用程序图标
        self.set_window_icon(self.root)

        # 初始化数据库
        self.init_db()

        # 导航历史记录（用于前进/后退）
        self.navigation_history = deque(maxlen=20)
        self.current_view = None

        # 初始视图 - 只显示项目列表
        self.show_project_panel()

        self.navigation_stack = []  # 使用栈来保存导航历史

    def set_window_icon(self, window):
        """统一设置窗口图标的方法"""
        try:
            # 尝试从当前目录加载图标
            if os.path.exists("application-note.ico"):
                window.iconbitmap("application-note.ico")
            else:
                # 如果不在当前目录，尝试在脚本所在目录查找
                script_dir = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(script_dir, "application-note.ico")
                if os.path.exists(icon_path):
                    window.iconbitmap(icon_path)
        except Exception as e:
            print(f"无法加载图标文件: {e}")

    def init_db(self):
        """初始化数据库"""
        self.conn = sqlite3.connect('project_manager.db')
        self.cursor = self.conn.cursor()
        self.cursor.execute('PRAGMA foreign_keys = ON')

        # 创建项目表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # 创建阶段表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS stages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                start_time TIMESTAMP,
                end_time TIMESTAMP,
                status TEXT DEFAULT '进行中',
                FOREIGN KEY (project_id) REFERENCES projects (id) ON DELETE CASCADE
            )
        ''')

        # 创建任务表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                stage_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (stage_id) REFERENCES stages (id) ON DELETE CASCADE
            )
        ''')

        # 创建项目文件表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS project_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects (id) ON DELETE CASCADE
            )
        ''')

        # 检查并添加缺失的列
        self._add_missing_columns()

        self.conn.commit()

    def _add_missing_columns(self):
        """添加缺失的列到现有表"""
        try:
            # 检查project_files表是否有file_size列
            self.cursor.execute("PRAGMA table_info(project_files)")
            columns = [column[1] for column in self.cursor.fetchall()]

            # 添加file_size列（如果不存在）
            if 'file_size' not in columns:
                self.cursor.execute('ALTER TABLE project_files ADD COLUMN file_size INTEGER')
                print("已添加file_size列到project_files表")

            # 添加file_type列（如果不存在）
            if 'file_type' not in columns:
                self.cursor.execute('ALTER TABLE project_files ADD COLUMN file_type TEXT')
                print("已添加file_type列到project_files表")

        except Exception as e:
            print(f"添加列时出错: {e}")

    def create_project_panel(self):
        """创建项目列表面板"""
        # 清除现有控件
        for widget in self.root.winfo_children():
            widget.destroy()

        # 顶部搜索框和按钮
        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=5, padx=5, fill=tk.X)

        ttk.Label(top_frame, text="搜索:").grid(row=0, column=0, padx=2)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=20)
        search_entry.grid(row=0, column=1, padx=2)
        search_entry.bind('<KeyRelease>', self.search_projects)

        search_btn = ttk.Button(top_frame, text="按日期搜索", command=self.search_by_date)
        search_btn.grid(row=0, column=2, padx=2)

        add_btn = ttk.Button(top_frame, text="新增项目", command=self.add_project)
        add_btn.grid(row=0, column=3, padx=2)

        # 新增导出按钮，改为下拉菜单
        export_menu_btn = ttk.Menubutton(top_frame, text="导出数据")
        export_menu = tk.Menu(export_menu_btn, tearoff=0)
        export_menu.add_command(label="导出Excel数据", command=self.export_to_excel)
        export_menu.add_command(label="导出完整项目（含文件）", command=self.export_full_project)
        export_menu_btn['menu'] = export_menu
        export_menu_btn.grid(row=0, column=4, padx=2)

        # 项目列表
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)

        # 修改列定义，将序号列放在第一列，项目名称放在第二列
        self.project_tree = ttk.Treeview(tree_frame, columns=('index', 'name', 'progress', 'created_at'),
                                         selectmode='browse')
        self.project_tree.heading('#0', text='')  # 隐藏默认的#0列
        self.project_tree.heading('index', text='序号')
        self.project_tree.heading('name', text='项目名称')  # 将项目名称作为单独的列
        self.project_tree.heading('progress', text='进度')
        self.project_tree.heading('created_at', text='创建时间')

        # 设置列宽
        self.project_tree.column('#0', width=0, stretch=False)  # 隐藏默认列
        self.project_tree.column('index', width=40, stretch=False)
        self.project_tree.column('name', width=200)  # 项目名称列宽
        self.project_tree.column('progress', width=80)
        self.project_tree.column('created_at', width=120)

        # 隐藏索引列（我们会在显示时填充它）
        self.project_tree.column('index', stretch=False)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.project_tree.yview)
        self.project_tree.configure(yscrollcommand=scrollbar.set)

        self.project_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定双击事件
        self.project_tree.bind('<Double-1>', self.on_project_double_click)

        # 创建完界面后加载数据
        self.load_projects()

    def export_full_project(self):
        """导出完整项目数据，支持多选项目到一个压缩包"""
        # 获取项目列表用于选择
        self.cursor.execute('''
            SELECT p.id, p.name 
            FROM projects p
            ORDER BY p.updated_at DESC
        ''')
        projects = self.cursor.fetchall()

        if not projects:
            messagebox.showinfo("提示", "没有可导出的项目")
            return

        # 创建选择对话框
        select_window = tk.Toplevel(self.root)
        select_window.title("选择导出项目")
        select_window.geometry("500x400")

        # 设置图标
        self.set_window_icon(select_window)

        tk.Label(select_window, text="请选择要导出的项目（可多选）：", font=('Arial', 12)).pack(pady=10)

        # 创建列表框（支持多选）
        listbox = tk.Listbox(select_window, width=50, height=15, selectmode=tk.MULTIPLE)
        listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 创建滚动条
        scrollbar = ttk.Scrollbar(select_window, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)

        # 填充项目列表
        project_dict = {}
        for project in projects:
            project_id, project_name = project
            display_text = f"{project_id:03d} - {project_name}"
            listbox.insert(tk.END, display_text)
            project_dict[display_text] = (project_id, project_name)

        # 创建底部按钮框架
        btn_frame = ttk.Frame(select_window)
        btn_frame.pack(pady=10, fill=tk.X)

        # 全选按钮
        def select_all():
            listbox.select_set(0, tk.END)

        # 取消全选按钮
        def deselect_all():
            listbox.select_clear(0, tk.END)

        # 导出按钮函数
        def on_export():
            selection_indices = listbox.curselection()
            if not selection_indices:
                messagebox.showwarning("警告", "请至少选择一个项目")
                return

            selected_projects = []
            for index in selection_indices:
                selected_text = listbox.get(index)
                project_id, project_name = project_dict[selected_text]
                selected_projects.append((project_id, project_name))

            # 关闭选择窗口
            select_window.destroy()

            # 根据选择的数量决定默认文件名
            if len(selected_projects) == 1:
                default_name = f"{selected_projects[0][1]}_完整数据.zip"
            else:
                default_name = f"项目数据_{len(selected_projects)}个项目.zip"

            # 让用户选择保存位置
            file_path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                filetypes=[("ZIP压缩包", "*.zip"), ("所有文件", "*.*")],
                title=f"保存项目数据（{len(selected_projects)}个项目）",
                initialfile=default_name
            )

            if not file_path:  # 用户取消了保存
                return

            try:
                # 创建临时目录
                temp_dir = os.path.join(os.path.dirname(file_path),
                                        f"temp_export_{datetime.now().strftime('%Y%m%d%H%M%S')}")
                os.makedirs(temp_dir, exist_ok=True)

                # 导出每个选中的项目
                all_files_to_zip = []  # 记录所有要添加到压缩包的文件

                for i, (project_id, project_name) in enumerate(selected_projects, 1):
                    print(f"正在导出项目 {i}/{len(selected_projects)}: {project_name}")

                    # 为每个项目创建子目录
                    project_dir = os.path.join(temp_dir, f"{project_id:03d}_{project_name}")
                    os.makedirs(project_dir, exist_ok=True)

                    # 1. 导出Excel数据
                    excel_file = os.path.join(project_dir, f"{project_name}_数据.xlsx")
                    self._export_project_to_excel_with_files(project_id, excel_file)
                    all_files_to_zip.append(
                        (excel_file, os.path.join(f"项目{project_id:03d}_{project_name}", f"{project_name}_数据.xlsx")))

                    # 2. 复制项目文件
                    project_files_dir = os.path.join(project_dir, "项目文件")
                    self._export_project_files_to_dir(project_id, project_files_dir)

                    # 记录项目文件
                    if os.path.exists(project_files_dir):
                        for root, dirs, files in os.walk(project_files_dir):
                            for file in files:
                                file_path_full = os.path.join(root, file)
                                rel_path = os.path.relpath(file_path_full, project_dir)
                                all_files_to_zip.append(
                                    (file_path_full, os.path.join(f"项目{project_id:03d}_{project_name}", rel_path)))

                # 3. 创建ZIP压缩包
                with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file_path_full, arcname in all_files_to_zip:
                        zipf.write(file_path_full, arcname=arcname)
                        print(f"  添加到压缩包: {arcname}")

                # 4. 清理临时目录
                shutil.rmtree(temp_dir)

                messagebox.showinfo("成功",
                                    f"已成功导出 {len(selected_projects)} 个项目到:\n{file_path}\n\n包含：\n1. Excel数据文件\n2. 所有上传的文件")

                # 询问是否打开导出目录
                if messagebox.askyesno("打开文件夹", "是否打开包含导出文件的文件夹？"):
                    self.open_file_in_explorer(os.path.dirname(file_path))

            except Exception as e:
                messagebox.showerror("错误", f"导出项目时出错:\n{str(e)}")
                # 清理临时目录
                if os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass

        # 按钮
        ttk.Button(btn_frame, text="全选", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消全选", command=deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出选中项目", command=on_export).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=select_window.destroy).pack(side=tk.LEFT, padx=5)

    def _export_project_files_to_dir(self, project_id, target_dir):
        """导出项目文件到指定目录"""
        # 获取项目文件
        self.cursor.execute('''
            SELECT file_path 
            FROM project_files 
            WHERE project_id = ?
        ''', (project_id,))
        files = self.cursor.fetchall()

        if files:
            os.makedirs(target_dir, exist_ok=True)

            # 复制文件
            for file_data in files:
                file_path = file_data[0]
                if os.path.exists(file_path):
                    try:
                        # 只复制文件名，避免路径冲突
                        file_name = os.path.basename(file_path)
                        dest_path = os.path.join(target_dir, file_name)

                        # 如果文件已存在，添加数字后缀
                        counter = 1
                        base_name, ext = os.path.splitext(file_name)
                        while os.path.exists(dest_path):
                            dest_path = os.path.join(target_dir, f"{base_name}_{counter}{ext}")
                            counter += 1

                        shutil.copy2(file_path, dest_path)
                    except Exception as e:
                        print(f"复制文件 {file_path} 时出错: {e}")

    def _export_project_to_excel_with_files(self, project_id, excel_file_path):
        """导出指定项目的数据到Excel（包含文件信息）"""
        # 创建Excel工作簿
        workbook = xlsxwriter.Workbook(excel_file_path)

        # 1. 项目基本信息
        self._export_project_info_to_excel(workbook, project_id)

        # 2. 阶段信息
        self._export_project_stages_to_excel(workbook, project_id)

        # 3. 任务信息
        self._export_project_tasks_to_excel(workbook, project_id)

        # 4. 文件信息
        self._export_project_files_to_excel(workbook, project_id)

        workbook.close()

    def _export_project_info_to_excel(self, workbook, project_id):
        """导出项目基本信息到Excel"""
        worksheet = workbook.add_worksheet("项目信息")

        # 获取项目信息
        self.cursor.execute('''
            SELECT name, created_at, updated_at 
            FROM projects 
            WHERE id = ?
        ''', (project_id,))
        project = self.cursor.fetchone()

        if project:
            name, created_at, updated_at = project

            # 写入标题
            worksheet.write(0, 0, "项目名称：")
            worksheet.write(0, 1, name)

            worksheet.write(1, 0, "创建时间：")
            worksheet.write(1, 1, created_at)

            worksheet.write(2, 0, "更新时间：")
            worksheet.write(2, 1, updated_at)

            # 设置列宽
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 30)

    def _export_project_stages_to_excel(self, workbook, project_id):
        """导出项目阶段信息到Excel"""
        worksheet = workbook.add_worksheet("阶段信息")

        # 标题行
        headers = ["序号", "阶段名称", "开始时间", "结束时间", "状态", "持续时间"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取阶段数据
        self.cursor.execute('''
            SELECT name, start_time, end_time, status 
            FROM stages 
            WHERE project_id = ? 
            ORDER BY start_time
        ''', (project_id,))
        stages = self.cursor.fetchall()

        # 写入数据
        for row, stage in enumerate(stages, start=1):
            name, start_time, end_time, status = stage

            # 计算持续时间
            duration = ""
            if start_time and end_time:
                try:
                    start = datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S')
                    end = datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S')
                    days = (end - start).days
                    if days > 0:
                        duration = f"{days}天"
                    else:
                        hours = (end - start).seconds // 3600
                        duration = f"{hours}小时"
                except:
                    duration = ""

            worksheet.write(row, 0, row)  # 序号
            worksheet.write(row, 1, name)
            worksheet.write(row, 2, start_time if start_time else "")
            worksheet.write(row, 3, end_time if end_time else "")
            worksheet.write(row, 4, status)
            worksheet.write(row, 5, duration)

        # 设置列宽
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:D', 18)
        worksheet.set_column('E:E', 10)
        worksheet.set_column('F:F', 12)

    def _export_project_tasks_to_excel(self, workbook, project_id):
        """导出项目任务信息到Excel"""
        worksheet = workbook.add_worksheet("任务信息")

        # 标题行
        headers = ["序号", "任务名称", "描述", "所属阶段", "创建时间"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取任务数据
        self.cursor.execute('''
            SELECT t.name, t.description, t.created_at, s.name as stage_name
            FROM tasks t
            JOIN stages s ON t.stage_id = s.id
            WHERE s.project_id = ?
            ORDER BY t.created_at
        ''', (project_id,))
        tasks = self.cursor.fetchall()

        # 写入数据
        for row, task in enumerate(tasks, start=1):
            name, description, created_at, stage_name = task
            worksheet.write(row, 0, row)  # 序号
            worksheet.write(row, 1, name)
            worksheet.write(row, 2, description if description else "")
            worksheet.write(row, 3, stage_name)
            worksheet.write(row, 4, created_at)

        # 设置列宽
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 40)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 18)

    def _export_project_files_to_excel(self, workbook, project_id):
        """导出项目文件信息到Excel"""
        worksheet = workbook.add_worksheet("文件信息")

        # 标题行
        headers = ["序号", "文件名称", "文件类型", "文件大小", "上传时间", "存储路径"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取文件数据
        self.cursor.execute('''
            SELECT file_name, file_path, uploaded_at, file_size, file_type
            FROM project_files 
            WHERE project_id = ? 
            ORDER BY uploaded_at DESC
        ''', (project_id,))
        files = self.cursor.fetchall()

        # 写入数据
        for row, file_data in enumerate(files, start=1):
            file_name, file_path, uploaded_at, file_size, file_type = file_data

            # 获取文件类型（扩展名）
            if not file_type and file_name:
                file_type = os.path.splitext(file_name)[1]

            # 格式化文件大小
            size_str = ""
            if file_size:
                if file_size < 1024:
                    size_str = f"{file_size} B"
                elif file_size < 1024 * 1024:
                    size_str = f"{file_size / 1024:.1f} KB"
                else:
                    size_str = f"{file_size / (1024 * 1024):.1f} MB"

            worksheet.write(row, 0, row)  # 序号
            worksheet.write(row, 1, file_name)
            worksheet.write(row, 2, file_type if file_type else "")
            worksheet.write(row, 3, size_str)
            worksheet.write(row, 4, uploaded_at)
            worksheet.write(row, 5, file_path)

        # 设置列宽
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 12)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:E', 18)
        worksheet.set_column('F:F', 50)

    def _export_project_files(self, project_id, project_name, temp_dir):
        """导出项目文件到临时目录"""
        # 获取项目文件
        self.cursor.execute('''
            SELECT file_path 
            FROM project_files 
            WHERE project_id = ?
        ''', (project_id,))
        files = self.cursor.fetchall()

        if files:
            # 创建项目文件目录
            project_files_dir = os.path.join(temp_dir, "项目文件", project_name)
            os.makedirs(project_files_dir, exist_ok=True)

            # 复制文件
            for file_data in files:
                file_path = file_data[0]
                if os.path.exists(file_path):
                    try:
                        shutil.copy2(file_path, project_files_dir)
                    except Exception as e:
                        print(f"复制文件 {file_path} 时出错: {e}")

    def open_file_in_explorer(self, path):
        """在文件资源管理器中打开指定路径"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # macOS, Linux
                subprocess.call(['open', path] if sys.platform == 'darwin' else ['xdg-open', path])
        except Exception as e:
            print(f"无法打开文件夹: {e}")

    def export_to_excel(self):
        """导出所有项目数据到Excel文件"""
        # 让用户选择保存位置
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            title="保存Excel文件",
            initialfile="项目数据.xlsx"
        )

        if not file_path:  # 用户取消了保存
            return

        try:
            # 创建Excel工作簿
            workbook = xlsxwriter.Workbook(file_path)

            # 添加项目工作表
            self._export_projects_to_excel(workbook)

            # 添加阶段工作表
            self._export_stages_to_excel(workbook)

            # 添加任务工作表
            self._export_tasks_to_excel(workbook)

            # 新增：添加文件工作表
            self._export_files_to_excel(workbook)

            workbook.close()
            messagebox.showinfo("成功", f"数据已成功导出到 {file_path}")

            # 询问是否打开文件
            if messagebox.askyesno("打开文件", "是否打开导出的Excel文件？"):
                self.open_excel_file(file_path)

        except Exception as e:
            messagebox.showerror("错误", f"导出Excel时出错:\n{str(e)}")

    def _export_files_to_excel(self, workbook):
        """导出所有文件数据到Excel"""
        worksheet = workbook.add_worksheet("文件列表")

        # 添加标题行
        headers = ["ID", "项目ID", "文件名称", "文件类型", "文件大小", "上传时间", "存储路径"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取文件数据
        self.cursor.execute('''
            SELECT pf.id, pf.project_id, pf.file_name, pf.file_type, 
                   pf.file_size, pf.uploaded_at, pf.file_path,
                   p.name as project_name
            FROM project_files pf
            JOIN projects p ON pf.project_id = p.id
            ORDER BY pf.project_id, pf.uploaded_at DESC
        ''')
        files = self.cursor.fetchall()

        # 写入数据
        for row, file_data in enumerate(files, start=1):
            file_id, project_id, file_name, file_type, file_size, uploaded_at, file_path, project_name = file_data

            # 获取文件类型（扩展名）
            if not file_type and file_name:
                file_type = os.path.splitext(file_name)[1]

            # 格式化文件大小
            size_str = ""
            if file_size:
                if file_size < 1024:
                    size_str = f"{file_size} B"
                elif file_size < 1024 * 1024:
                    size_str = f"{file_size / 1024:.1f} KB"
                else:
                    size_str = f"{file_size / (1024 * 1024):.1f} MB"

            worksheet.write(row, 0, file_id)
            worksheet.write(row, 1, f"{project_id} ({project_name})")
            worksheet.write(row, 2, file_name)
            worksheet.write(row, 3, file_type if file_type else "")
            worksheet.write(row, 4, size_str)
            worksheet.write(row, 5, uploaded_at)
            worksheet.write(row, 6, file_path)

        # 设置列宽
        worksheet.set_column('A:A', 8)  # ID
        worksheet.set_column('B:B', 15)  # 项目ID
        worksheet.set_column('C:C', 30)  # 文件名称
        worksheet.set_column('D:D', 12)  # 文件类型
        worksheet.set_column('E:E', 12)  # 文件大小
        worksheet.set_column('F:F', 18)  # 上传时间
        worksheet.set_column('G:G', 50)  # 存储路径

    def open_excel_file(self, file_path):
        """打开Excel文件"""
        try:
            if os.name == 'nt':  # Windows
                os.startfile(file_path)
            elif os.name == 'posix':  # macOS, Linux
                subprocess.call(['open', file_path] if sys.platform == 'darwin' else ['xdg-open', file_path])
        except Exception as e:
            messagebox.showwarning("警告", f"无法打开Excel文件:\n{str(e)}")

    def _export_projects_to_excel(self, workbook):
        """导出项目数据到Excel"""
        worksheet = workbook.add_worksheet("项目")

        # 添加标题行
        headers = ["ID", "项目名称", "创建时间", "更新时间", "总阶段数", "已完成阶段数", "进度"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取项目数据
        self.cursor.execute('''
            SELECT p.id, p.name, p.created_at, p.updated_at, 
                   COUNT(s.id) as total_stages,
                   SUM(CASE WHEN s.status = '已完成' THEN 1 ELSE 0 END) as completed_stages
            FROM projects p
            LEFT JOIN stages s ON p.id = s.project_id
            GROUP BY p.id
            ORDER BY p.updated_at DESC
        ''')
        projects = self.cursor.fetchall()

        # 写入数据
        for row, project in enumerate(projects, start=1):
            project_id, name, created_at, updated_at, total_stages, completed_stages = project

            # 计算进度
            progress = 0
            if total_stages > 0:
                progress = (completed_stages / total_stages) * 100

            worksheet.write(row, 0, project_id)
            worksheet.write(row, 1, name)
            worksheet.write(row, 2, created_at)
            worksheet.write(row, 3, updated_at)
            worksheet.write(row, 4, total_stages)
            worksheet.write(row, 5, completed_stages)
            worksheet.write(row, 6, f"{progress:.1f}%")

        # 设置列宽
        worksheet.set_column('A:A', 8)  # ID
        worksheet.set_column('B:B', 30)  # 项目名称
        worksheet.set_column('C:D', 18)  # 时间
        worksheet.set_column('E:F', 12)  # 阶段数
        worksheet.set_column('G:G', 10)  # 进度

    def _export_stages_to_excel(self, workbook):
        """导出阶段数据到Excel"""
        worksheet = workbook.add_worksheet("阶段")

        # 添加标题行
        headers = ["ID", "项目ID", "阶段名称", "开始时间", "结束时间", "状态"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取阶段数据
        self.cursor.execute('''
            SELECT id, project_id, name, start_time, end_time, status 
            FROM stages 
            ORDER BY project_id, start_time
        ''')
        stages = self.cursor.fetchall()

        # 写入数据
        for row, stage in enumerate(stages, start=1):
            for col, value in enumerate(stage):
                worksheet.write(row, col, value)

        # 设置列宽
        worksheet.set_column('A:A', 8)  # ID
        worksheet.set_column('B:B', 8)  # 项目ID
        worksheet.set_column('C:C', 30)  # 阶段名称
        worksheet.set_column('D:E', 18)  # 时间
        worksheet.set_column('F:F', 12)  # 状态

    def _export_tasks_to_excel(self, workbook):
        """导出任务数据到Excel"""
        worksheet = workbook.add_worksheet("任务")

        # 添加标题行
        headers = ["ID", "阶段ID", "任务名称", "描述", "创建时间"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # 获取任务数据
        self.cursor.execute('''
            SELECT id, stage_id, name, description, created_at 
            FROM tasks 
            ORDER BY stage_id, created_at
        ''')
        tasks = self.cursor.fetchall()

        # 写入数据
        for row, task in enumerate(tasks, start=1):
            for col, value in enumerate(task):
                # 处理可能为None的描述字段
                if col == 3 and value is None:
                    value = ""
                worksheet.write(row, col, value)

        # 设置列宽
        worksheet.set_column('A:A', 8)  # ID
        worksheet.set_column('B:B', 8)  # 阶段ID
        worksheet.set_column('C:C', 30)  # 任务名称
        worksheet.set_column('D:D', 50)  # 描述
        worksheet.set_column('E:E', 18)  # 创建时间

    def show_project_panel(self):
        """显示项目面板"""
        self.current_view = 'project_list'
        self.create_project_panel()  # 这个方法内部会调用 load_projects()

    def load_projects(self):
        """加载项目数据"""
        # 检查 project_tree 是否存在并且是否有效
        if not hasattr(self, 'project_tree'):
            return

        # 检查控件是否仍然存在
        try:
            if not self.project_tree.winfo_exists():
                return
        except:
            return  # 如果检查失败，说明控件已不存在

        # 清空现有数据
        for item in self.project_tree.get_children():
            self.project_tree.delete(item)

        # 从数据库获取项目数据
        self.cursor.execute('''
            SELECT p.id, p.name, p.created_at, 
                   COUNT(s.id) as total_stages,
                   SUM(CASE WHEN s.status = '已完成' THEN 1 ELSE 0 END) as completed_stages
            FROM projects p
            LEFT JOIN stages s ON p.id = s.project_id
            GROUP BY p.id
            ORDER BY p.updated_at DESC
        ''')

        projects = self.cursor.fetchall()

        for index, project in enumerate(projects, start=1):
            project_id, name, created_at, total_stages, completed_stages = project

            # 计算进度
            progress = 0
            if total_stages > 0:
                progress = (completed_stages / total_stages) * 100

            # 插入到树形视图，正确顺序：序号、项目名称、进度、创建时间
            item_id = self.project_tree.insert('', 'end',
                                               values=(index, name, f"{progress:.1f}%", created_at),
                                               tags=(str(project_id),))

    def navigate_to(self, view_name, *args):
        """通用导航方法"""
        # 保存当前状态到导航栈
        if hasattr(self, 'current_view'):
            self.navigation_stack.append((self.current_view, self.get_current_view_data()))

        # 导航到新视图
        if view_name == 'project_list':
            self.show_project_panel()
        elif view_name == 'project_detail':
            self.show_project_detail(*args)
        elif view_name == 'stage_detail':
            self.show_stage_detail(*args)
        elif view_name == 'task_detail':
            self.show_task_detail(*args)

    def get_current_view_data(self):
        """获取当前视图的状态数据，用于恢复"""
        if not hasattr(self, 'current_view'):
            return None

        if self.current_view == 'project_list':
            return None
        elif self.current_view[0] == 'project_detail':
            return (self.current_project_id, self.project_name_var.get())
        elif self.current_view[0] == 'stage_detail':
            return (self.current_project_id, self.project_name_var.get(),
                    self.current_stage_id, self.stage_name_var.get())
        elif self.current_view[0] == 'task_detail':
            # 对于任务详情，我们需要保存足够的信息回到阶段详情
            return (self.current_project_id, self.project_name_var.get(),
                    self.current_stage_id, self.stage_name_var.get(),
                    self.current_task_id, self.task_name_var.get(),
                    self.desc_text.get('1.0', tk.END).strip())

    def search_projects(self, event=None):
        """搜索项目"""
        search_term = self.search_var.get().lower()

        # 清空现有数据
        for item in self.project_tree.get_children():
            self.project_tree.delete(item)

        # 从数据库获取项目数据
        self.cursor.execute('''
            SELECT p.id, p.name, p.created_at, 
                   COUNT(s.id) as total_stages,
                   SUM(CASE WHEN s.status = '已完成' THEN 1 ELSE 0 END) as completed_stages
            FROM projects p
            LEFT JOIN stages s ON p.id = s.project_id
            WHERE LOWER(p.name) LIKE ?
            GROUP BY p.id
            ORDER BY p.updated_at DESC
        ''', (f'%{search_term}%',))

        projects = self.cursor.fetchall()

        for index, project in enumerate(projects, start=1):
            project_id, name, created_at, total_stages, completed_stages = project

            # 计算进度
            progress = 0
            if total_stages > 0:
                progress = (completed_stages / total_stages) * 100

            # 插入到树形视图，添加序号
            item_id = self.project_tree.insert('', 'end',
                                               values=(index, name, f"{progress:.1f}%", created_at),
                                               tags=(str(project_id),))

    def search_by_date(self):
        """按日期搜索项目"""
        date_str = simpledialog.askstring("按日期搜索", "请输入日期 (YYYY-MM-DD):")
        if date_str:
            try:
                # 验证日期格式
                datetime.strptime(date_str, '%Y-%m-%d')

                # 清空现有数据
                for item in self.project_tree.get_children():
                    self.project_tree.delete(item)

                # 从数据库获取项目数据
                self.cursor.execute('''
                    SELECT p.id, p.name, p.created_at, 
                           COUNT(s.id) as total_stages,
                           SUM(CASE WHEN s.status = '已完成' THEN 1 ELSE 0 END) as completed_stages
                    FROM projects p
                    LEFT JOIN stages s ON p.id = s.project_id
                    WHERE DATE(p.created_at) = ? OR DATE(p.updated_at) = ?
                    GROUP BY p.id
                    ORDER BY p.updated_at DESC
                ''', (date_str, date_str))

                projects = self.cursor.fetchall()

                for index, project in enumerate(projects, start=1):
                    project_id, name, created_at, total_stages, completed_stages = project

                    # 计算进度
                    progress = 0
                    if total_stages > 0:
                        progress = (completed_stages / total_stages) * 100

                    # 插入到树形视图，添加序号
                    item_id = self.project_tree.insert('', 'end',
                                                       values=(index, name, f"{progress:.1f}%", created_at),
                                                       tags=(str(project_id),))

            except ValueError:
                messagebox.showerror("错误", "日期格式不正确，请使用YYYY-MM-DD格式")

    def add_project(self):
        """添加新项目"""
        name = simpledialog.askstring("新增项目", "请输入项目名称:")
        if name:
            # 使用当前本地时间而不是数据库的CURRENT_TIMESTAMP
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute("INSERT INTO projects (name, created_at, updated_at) VALUES (?, ?, ?)",
                                (name, current_time, current_time))
            self.conn.commit()
            self.load_projects()

    def on_project_double_click(self, event):
        item = self.project_tree.selection()
        if not item:
            return

        item = item[0]
        project_id = self.project_tree.item(item, 'tags')[0]
        project_name = self.project_tree.item(item, 'values')[1]  # 从values中获取项目名称

        # 使用导航方法
        self.navigate_to('project_detail', project_id, project_name)

    def show_project_detail(self, project_id, project_name):
        """显示项目详情面板 - 分为左右视图"""
        # 清除现有控件
        for widget in self.root.winfo_children():
            widget.destroy()

        self.current_view = ('project_detail', project_id)
        self.current_project_id = project_id

        # 创建主界面面板容器 - 左右分割
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        # 左侧面板（阶段视图）
        self.left_panel = ttk.Frame(self.main_paned, width=600)
        self.main_paned.add(self.left_panel, weight=2)

        # 右侧面板（文件视图）
        self.right_panel = ttk.Frame(self.main_paned, width=400)
        self.main_paned.add(self.right_panel, weight=2)

        # 顶部工具栏
        toolbar = ttk.Frame(self.left_panel)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        back_btn = ttk.Button(toolbar, text="← 返回项目列表", command=self.navigate_back)
        back_btn.pack(side=tk.LEFT, padx=2)

        ttk.Label(toolbar, text=f"项目: {project_name}", font=('Arial', 12, 'bold')).pack(side=tk.LEFT, padx=10)

        # 项目信息区域
        info_frame = ttk.LabelFrame(self.left_panel, text="项目信息", padding=10)
        info_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(info_frame, text="项目名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.project_name_var = tk.StringVar(value=project_name)
        name_entry = ttk.Entry(info_frame, textvariable=self.project_name_var, width=40)
        name_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # 按钮放在第二行
        button_frame = ttk.Frame(info_frame)
        button_frame.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=5)

        save_btn = ttk.Button(button_frame, text="保存", command=self.save_project_name)
        save_btn.pack(side=tk.LEFT, padx=5)

        add_stage_btn = ttk.Button(button_frame, text="新增阶段",
                                   command=lambda: self.add_stage(project_id))
        add_stage_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(button_frame, text="删除项目",
                                command=lambda: self.delete_project(project_id))
        delete_btn.pack(side=tk.LEFT, padx=5)

        # 上传文件按钮 - 放在右侧面板的顶部
        right_toolbar = ttk.Frame(self.right_panel)
        right_toolbar.pack(fill=tk.X, padx=5, pady=5)

        upload_btn = ttk.Button(right_toolbar, text="上传文件",
                                command=lambda: self.upload_file(project_id, project_name))
        upload_btn.pack(side=tk.LEFT, padx=2)

        # 新增删除文件按钮
        delete_file_btn = ttk.Button(right_toolbar, text="删除文件",
                                     command=lambda: self.delete_project_file(project_id))
        delete_file_btn.pack(side=tk.LEFT, padx=2)

        # 阶段视图区域
        stage_frame = ttk.LabelFrame(self.left_panel, text="项目阶段", padding=10)
        stage_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建Canvas用于显示阶段
        self.stage_canvas_frame = ttk.Frame(stage_frame)
        self.stage_canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.stage_canvas = tk.Canvas(self.stage_canvas_frame, bg='white')

        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(self.stage_canvas_frame, orient=tk.VERTICAL, command=self.stage_canvas.yview)
        h_scrollbar = ttk.Scrollbar(self.stage_canvas_frame, orient=tk.HORIZONTAL, command=self.stage_canvas.xview)

        self.stage_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # 布局
        self.stage_canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        self.stage_canvas_frame.grid_rowconfigure(0, weight=1)
        self.stage_canvas_frame.grid_columnconfigure(0, weight=1)

        # 绑定Canvas调整大小事件
        self.stage_canvas.bind('<Configure>', self.on_stage_canvas_configure)

        # 绑定Canvas右键事件（新增）
        self.stage_canvas.bind('<Button-3>', self.on_stage_right_click)
        self.stage_canvas.bind('<Button-1>', self.on_stage_canvas_click)

        # 文件列表区域 - 在右侧面板
        file_frame = ttk.LabelFrame(self.right_panel, text="项目文件", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建文件列表树形视图
        columns = ('file_id', 'uploaded_at', 'file_size')
        self.file_tree = ttk.Treeview(file_frame, columns=columns, selectmode='browse', show='tree headings')
        self.file_tree.heading('#0', text='文件名称')
        self.file_tree.heading('uploaded_at', text='上传时间')
        self.file_tree.heading('file_size', text='文件大小')

        # 隐藏额外的列
        self.file_tree.column('file_id', width=0, stretch=False)

        self.file_tree.column('#0', width=200)
        self.file_tree.column('uploaded_at', width=120)
        self.file_tree.column('file_size', width=80)

        # 添加滚动条
        file_scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=file_scrollbar.set)

        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定双击事件
        self.file_tree.bind('<Double-1>', self.on_file_double_click)

        # 加载阶段数据
        self.load_stages(project_id)
        self.load_project_files(project_id)

    def on_stage_right_click(self, event):
        """处理阶段右键点击，显示删除菜单"""
        # 检查 Canvas 是否仍然存在
        if not hasattr(self, 'stage_canvas') or not self.stage_canvas.winfo_exists():
            return

        # 检查点击的是哪个阶段
        clicked_stage_id = None
        for stage_id, data in self.stage_rects.items():
            x1, y1, x2, y2 = data['coords']
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                clicked_stage_id = stage_id
                break

        if clicked_stage_id:
            # 创建右键菜单
            menu = tk.Menu(self.root, tearoff=0)
            stage_name = self.stage_rects[clicked_stage_id]['name']
            menu.add_command(label=f"删除阶段: {stage_name}",
                             command=lambda s_id=clicked_stage_id: self.delete_stage(s_id))

            try:
                menu.post(event.x_root, event.y_root)
            except Exception as e:
                # 如果发布菜单失败，静默处理
                pass

            # 保存点击的阶段ID
            self.right_clicked_stage_id = clicked_stage_id

    def on_stage_canvas_click(self, event):
        """点击Canvas其他地方时处理"""
        if hasattr(self, 'right_clicked_stage_id'):
            self.right_clicked_stage_id = None

    def delete_stage(self, stage_id):
        """删除阶段及其所有任务"""
        if not stage_id:
            return

        # 获取阶段信息
        stage_name = ""
        for stage in self.stages:
            if stage[0] == stage_id:
                stage_name = stage[1]
                break

        if not stage_name:
            return

        # 确认删除
        if not messagebox.askyesno("确认删除", f"确定要删除阶段 '{stage_name}' 吗？\n这将删除该阶段的所有任务！"):
            return

        try:
            # 删除阶段（由于外键约束，相关的任务会自动删除）
            self.cursor.execute("DELETE FROM stages WHERE id = ?", (stage_id,))

            # 更新项目更新时间
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute("UPDATE projects SET updated_at = ? WHERE id = ?",
                                (current_time, self.current_project_id))

            self.conn.commit()

            # 刷新阶段视图 - 使用安全方式
            if hasattr(self, 'stage_canvas') and self.stage_canvas.winfo_exists():
                self.load_stages(self.current_project_id)

            # 刷新项目列表中的进度 - 使用安全方式
            if hasattr(self, 'project_tree') and self.project_tree.winfo_exists():
                self.load_projects()

            messagebox.showinfo("成功", f"阶段 '{stage_name}' 已删除")

        except Exception as e:
            messagebox.showerror("错误", f"删除阶段时出错:\n{str(e)}")

    def upload_file(self, project_id, project_name):
        """上传文件到项目"""
        file_path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[("所有文件", "*.*")]
        )

        if not file_path:
            return

        try:
            # 创建项目文件夹
            project_folder = os.path.join(os.getcwd(), "项目文件", project_name)
            os.makedirs(project_folder, exist_ok=True)

            # 复制文件到项目文件夹
            file_name = os.path.basename(file_path)
            dest_path = os.path.join(project_folder, file_name)

            # 如果文件已存在，询问是否覆盖
            if os.path.exists(dest_path):
                if not messagebox.askyesno("文件已存在", f"文件 '{file_name}' 已存在，是否覆盖？"):
                    return

            shutil.copy2(file_path, dest_path)

            # 获取文件信息
            file_size = os.path.getsize(dest_path)
            file_type = os.path.splitext(file_name)[1]

            # 记录文件信息到数据库
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute(
                "INSERT INTO project_files (project_id, file_name, file_path, file_size, file_type, uploaded_at) VALUES (?, ?, ?, ?, ?, ?)",
                (project_id, file_name, dest_path, file_size, file_type, current_time)
            )
            self.conn.commit()

            # 刷新文件列表
            self.load_project_files(project_id)

            messagebox.showinfo("成功", f"文件 '{file_name}' 上传成功")

        except Exception as e:
            messagebox.showerror("错误", f"上传文件时出错: {str(e)}")

    def load_project_files(self, project_id):
        """加载项目文件列表"""
        # 检查 file_tree 是否存在
        if not hasattr(self, 'file_tree'):
            return

        # 清空现有数据
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # 从数据库获取文件数据
        self.cursor.execute('''
            SELECT id, file_name, uploaded_at, file_size 
            FROM project_files 
            WHERE project_id = ? 
            ORDER BY uploaded_at DESC
        ''', (project_id,))

        files = self.cursor.fetchall()

        for file in files:
            file_id, file_name, uploaded_at, file_size = file
            # 简化显示时间
            display_time = uploaded_at.split()[0] if uploaded_at else ""

            # 格式化文件大小
            size_str = ""
            if file_size:
                if file_size < 1024:
                    size_str = f"{file_size} B"
                elif file_size < 1024 * 1024:
                    size_str = f"{file_size / 1024:.1f} KB"
                else:
                    size_str = f"{file_size / (1024 * 1024):.1f} MB"

            self.file_tree.insert('', 'end', text=file_name,
                                  values=(file_id, display_time, size_str))

    def delete_project_file(self, project_id):
        """删除项目文件"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的文件")
            return

        item = selected_items[0]
        file_id = self.file_tree.item(item, 'values')[0]
        file_name = self.file_tree.item(item, 'text')

        if not messagebox.askyesno("确认删除", f"确定要删除文件 '{file_name}' 吗？"):
            return

        try:
            # 从数据库获取文件路径
            self.cursor.execute('SELECT file_path FROM project_files WHERE id = ?', (file_id,))
            result = self.cursor.fetchone()

            if result:
                file_path = result[0]

                # 删除数据库记录
                self.cursor.execute('DELETE FROM project_files WHERE id = ?', (file_id,))
                self.conn.commit()

                # 删除物理文件
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        messagebox.showwarning("警告", f"文件已从数据库删除，但物理文件删除失败:\n{str(e)}")

                # 刷新文件列表
                self.load_project_files(project_id)
                messagebox.showinfo("成功", f"文件 '{file_name}' 已删除")

        except Exception as e:
            messagebox.showerror("错误", f"删除文件时出错: {str(e)}")

    def on_file_double_click(self, event):
        """双击打开文件"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            return

        item = selected_items[0]
        file_id = self.file_tree.item(item, 'values')[0]
        file_name = self.file_tree.item(item, 'text')

        # 从数据库获取文件路径
        self.cursor.execute('SELECT file_path FROM project_files WHERE id = ?', (file_id,))
        result = self.cursor.fetchone()

        if result:
            file_path = result[0]
            try:
                # 使用系统默认程序打开文件
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS, Linux
                    subprocess.call(('open', file_path) if sys.platform == 'darwin' else ('xdg-open', file_path))
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {str(e)}")

    def navigate_back(self):
        """返回上一视图"""
        if not self.navigation_stack:
            return

        # 检查是否有未保存的修改
        if self.current_view and self.current_view[0] == 'task_detail' and hasattr(self,
                                                                                   'changes_made') and self.changes_made:
            response = messagebox.askyesnocancel("警告", "刚刚修改的内容还没有保存，是否保存？",
                                                 default=messagebox.YES)
            if response is None:  # 用户点击了取消
                return
            elif response:  # 用户点击了是
                self.save_task()
            # 如果用户点击否，则继续返回

        # 弹出上一个视图状态
        prev_view, prev_data = self.navigation_stack.pop()

        # 恢复上一个视图
        if prev_view == 'project_list':
            self.show_project_panel()
        elif prev_view[0] == 'project_detail':
            self.show_project_detail(*prev_data)
        elif prev_view[0] == 'stage_detail':
            self.show_stage_detail(prev_data[2], prev_data[3])  # stage_id, stage_name
        elif prev_view[0] == 'task_detail':
            # 理论上不应该直接跳回任务详情，而是回到阶段详情
            self.show_stage_detail(prev_data[2], prev_data[3])  # stage_id, stage_name

    def save_project_name(self):
        """保存项目名称修改"""
        new_name = self.project_name_var.get()
        if new_name:
            # 使用当前本地时间而不是数据库的CURRENT_TIMESTAMP
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute("UPDATE projects SET name = ?, updated_at = ? WHERE id = ?",
                                (new_name, current_time, self.current_project_id))
            self.conn.commit()
            # 刷新项目列表
            self.load_projects()

    def delete_project(self, project_id):
        """删除项目"""
        if messagebox.askyesno("确认删除", "确定要删除这个项目吗？这将删除所有相关的阶段、任务和文件。"):
            try:
                # 获取项目名称
                self.cursor.execute("SELECT name FROM projects WHERE id = ?", (project_id,))
                project_name = self.cursor.fetchone()[0]

                # 先删除相关的任务
                self.cursor.execute("DELETE FROM tasks WHERE stage_id IN (SELECT id FROM stages WHERE project_id = ?)",
                                    (project_id,))
                # 删除相关的阶段
                self.cursor.execute("DELETE FROM stages WHERE project_id = ?", (project_id,))
                # 删除相关的文件记录
                self.cursor.execute("DELETE FROM project_files WHERE project_id = ?", (project_id,))
                # 删除项目
                self.cursor.execute("DELETE FROM projects WHERE id = ?", (project_id,))
                self.conn.commit()

                # 尝试删除项目文件夹
                project_folder = os.path.join(os.getcwd(), "项目文件", project_name)
                if os.path.exists(project_folder):
                    try:
                        shutil.rmtree(project_folder)
                    except:
                        pass  # 如果文件夹删除失败，不影响主流程

                # 返回项目列表
                self.show_project_panel()

            except Exception as e:
                messagebox.showerror("错误", f"删除项目时出错: {str(e)}")

    def add_stage(self, project_id):
        """添加新阶段"""
        name = simpledialog.askstring("新增阶段", "请输入阶段名称:")
        if name:
            # 获取当前时间
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 如果有其他阶段，结束上一个阶段
            self.cursor.execute("SELECT id FROM stages WHERE project_id = ? ORDER BY start_time DESC LIMIT 1",
                                (project_id,))
            last_stage = self.cursor.fetchone()
            if last_stage:
                self.cursor.execute("UPDATE stages SET end_time = ?, status = '已完成' WHERE id = ?",
                                    (current_time, last_stage[0]))

            # 添加新阶段
            self.cursor.execute(
                "INSERT INTO stages (project_id, name, start_time, status) VALUES (?, ?, ?, '进行中')",
                (project_id, name, current_time))

            # 更新项目更新时间
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute("UPDATE projects SET updated_at = ? WHERE id = ?",
                                (current_time, project_id))

            self.conn.commit()
            self.load_stages(project_id)
            # 刷新项目列表中的进度
            self.load_projects()

    def load_stages(self, project_id):
        """加载阶段数据"""
        # 检查stage_canvas是否存在，如果不存在则直接返回
        if not hasattr(self, 'stage_canvas') or not self.stage_canvas.winfo_exists():
            return

        # 清空Canvas
        self.stage_canvas.delete('all')

        # 从数据库获取阶段数据
        self.cursor.execute('''
            SELECT id, name, start_time, end_time, status 
            FROM stages 
            WHERE project_id = ? 
            ORDER BY start_time
        ''', (project_id,))

        self.stages = self.cursor.fetchall()

        if not self.stages:
            return

        # 计算布局
        self.layout_stages()

    def layout_stages(self):
        """布局阶段方块 - 修正蛇形布局"""
        stage_width = 150
        stage_height = 100
        h_spacing = 100
        v_spacing = 120

        # 清空Canvas并重置矩形映射
        self.stage_canvas.delete('all')
        self.stage_rects = {}
        self.stage_info = {}

        if not self.stages:
            return

        # 计算Canvas可用宽度
        canvas_width = self.stage_canvas.winfo_width() - 20
        if canvas_width < stage_width + h_spacing:
            canvas_width = stage_width + h_spacing + 20

        # 计算每行可以放置的阶段数量
        stages_per_row = max(1, (canvas_width - 20) // (stage_width + h_spacing))

        # 分组阶段到行 - 保持原始顺序
        rows = []
        for i in range(0, len(self.stages), stages_per_row):
            row = self.stages[i:i + stages_per_row]
            rows.append(row)

        # 先绘制所有阶段方块并存储坐标
        for row_idx, row in enumerate(rows):
            y = 50 + row_idx * v_spacing
            is_right_to_left = (row_idx % 2 == 1)

            for col_idx, stage in enumerate(row):
                stage_id, name, start_time, end_time, status = stage

                # 计算x坐标
                if is_right_to_left:
                    x = canvas_width - 20 - stage_width - col_idx * (stage_width + h_spacing)
                else:
                    x = 20 + col_idx * (stage_width + h_spacing)

                # 绘制阶段方块
                fill_color = 'lightgreen' if status == '已完成' else 'lightblue'
                rect = self.stage_canvas.create_rectangle(
                    x, y, x + stage_width, y + stage_height,
                    fill=fill_color, outline='black', width=2
                )

                # 存储矩形坐标和阶段信息
                self.stage_rects[stage_id] = {
                    'rect': rect,
                    'coords': (x, y, x + stage_width, y + stage_height),
                    'name': name,
                    'start_time': start_time,
                    'end_time': end_time,
                    'status': status,
                    'row': row_idx,
                    'col': col_idx,
                    'visual_col': col_idx if not is_right_to_left else (len(row) - 1 - col_idx)
                }

                # 存储阶段信息用于箭头连接
                self.stage_info[stage_id] = {
                    'id': stage_id,
                    'name': name,
                    'logical_index': self.stages.index(stage),
                    'visual_row': row_idx,
                    'visual_col': col_idx if not is_right_to_left else (len(row) - 1 - col_idx),
                    'rect_coords': (x, y, x + stage_width, y + stage_height)
                }

                # 添加阶段名称
                self.stage_canvas.create_text(
                    x + stage_width / 2, y + 20,
                    text=name, font=('Arial', 12, 'bold'),
                    width=stage_width - 10
                )

                # 添加时间信息
                time_text = f"开始: {start_time.split()[0] if start_time else '未设置'}"
                if end_time:
                    time_text += f"\n结束: {end_time.split()[0]}"

                self.stage_canvas.create_text(
                    x + stage_width / 2, y + stage_height / 2,
                    text=time_text, font=('Arial', 10),
                    width=stage_width - 10
                )

                # 添加状态
                self.stage_canvas.create_text(
                    x + stage_width / 2, y + stage_height - 20,
                    text=f"状态: {status}", font=('Arial', 10)
                )

                # 绑定双击事件
                self.stage_canvas.tag_bind(rect, '<Double-1>',
                                           lambda e, s_id=stage_id: self.on_stage_double_click(s_id))

        # 绘制箭头 - 按照逻辑顺序连接
        sorted_stages = sorted(self.stages, key=lambda x: x[0])

        for i in range(len(sorted_stages) - 1):
            current_stage = sorted_stages[i]
            next_stage = sorted_stages[i + 1]

            current_id = current_stage[0]
            next_id = next_stage[0]

            if current_id not in self.stage_info or next_id not in self.stage_info:
                continue

            current = self.stage_info[current_id]
            next_stage_info = self.stage_info[next_id]

            x1, y1, x1_end, y1_end = current['rect_coords']
            x2, y2, x2_end, y2_end = next_stage_info['rect_coords']

            current_center_x = (x1 + x1_end) / 2
            current_center_y = (y1 + y1_end) / 2
            next_center_x = (x2 + x2_end) / 2
            next_center_y = (y2 + y2_end) / 2

            same_row = current['visual_row'] == next_stage_info['visual_row']

            if same_row:
                if current['visual_col'] < next_stage_info['visual_col']:
                    start_x = x1_end
                    end_x = x2
                    mid_y = current_center_y
                else:
                    start_x = x1
                    end_x = x2_end
                    mid_y = current_center_y

                self.stage_canvas.create_line(
                    start_x, mid_y,
                    end_x, mid_y,
                    arrow=tk.LAST, width=2, smooth=True
                )
            else:
                if current['visual_col'] == 0:
                    start_x = x1
                else:
                    start_x = x1_end

                if next_stage_info['visual_col'] == 0:
                    end_x = x2
                else:
                    end_x = x2_end

                mid_x1 = start_x
                mid_y1 = current_center_y
                mid_x2 = end_x
                mid_y2 = next_center_y

                if abs(current['visual_row'] - next_stage_info['visual_row']) > 1:
                    canvas_center_x = canvas_width / 2
                    self.stage_canvas.create_line(
                        start_x, mid_y1,
                        canvas_center_x, mid_y1,
                        canvas_center_x, mid_y2,
                        end_x, mid_y2,
                        arrow=tk.LAST, width=2, smooth=True
                    )
                else:
                    self.stage_canvas.create_line(
                        start_x, mid_y1,
                        start_x, (mid_y1 + mid_y2) / 2,
                        end_x, (mid_y1 + mid_y2) / 2,
                        end_x, mid_y2,
                        arrow=tk.LAST, width=2, smooth=True
                    )

        # 更新滚动区域
        total_height = 50 + len(rows) * v_spacing
        self.stage_canvas.configure(scrollregion=(0, 0, canvas_width, total_height))

    def on_stage_canvas_configure(self, event):
        """Canvas大小改变时重新布局"""
        if hasattr(self, 'stages'):
            self.layout_stages()

    def on_stage_double_click(self, stage_id):
        stage_data = None
        for stage in self.stages:
            if stage[0] == stage_id:
                stage_data = stage
                break

        if stage_data:
            stage_name = stage_data[1]
            self.navigate_to('stage_detail', stage_id, stage_name)

    def show_stage_detail(self, stage_id, stage_name):
        """显示阶段详情面板"""
        # 清除现有控件
        for widget in self.root.winfo_children():
            widget.destroy()

        self.current_view = ('stage_detail', stage_id)

        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        back_btn = ttk.Button(toolbar, text="← 返回项目详情", command=self.navigate_back)
        back_btn.pack(side=tk.LEFT, padx=2)

        ttk.Label(toolbar, text=f"阶段: {stage_name}", font=('Arial', 12, 'bold')).pack(side=tk.LEFT, padx=10)

        # 主内容区域
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 阶段信息区域
        info_frame = ttk.LabelFrame(main_frame, text="阶段信息", padding=10)
        info_frame.pack(fill=tk.X, pady=5)

        # 获取阶段数据
        self.cursor.execute('''
            SELECT name, start_time, end_time, status 
            FROM stages 
            WHERE id = ?
        ''', (stage_id,))
        stage_data = self.cursor.fetchone()
        original_stage_name, start_time, end_time, status = stage_data

        ttk.Label(info_frame, text="阶段名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.stage_name_var = tk.StringVar(value=original_stage_name)
        name_entry = ttk.Entry(info_frame, textvariable=self.stage_name_var, width=40)
        name_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # 按钮放在第二行
        button_frame = ttk.Frame(info_frame)
        button_frame.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=5)

        save_btn = ttk.Button(button_frame, text="保存", command=self.save_stage_name)
        save_btn.pack(side=tk.LEFT, padx=5)

        complete_btn = ttk.Button(button_frame, text="结束阶段", command=self.complete_stage)
        complete_btn.pack(side=tk.LEFT, padx=5)

        # 新增删除阶段按钮
        delete_stage_btn = ttk.Button(button_frame, text="删除阶段",
                                      command=lambda: self.delete_stage_from_detail(stage_id, stage_name))
        delete_stage_btn.pack(side=tk.LEFT, padx=5)

        add_task_btn = ttk.Button(button_frame, text="新增任务",
                                  command=lambda: self.add_task(stage_id))
        add_task_btn.pack(side=tk.LEFT, padx=5)

        # 时间信息
        time_frame = ttk.Frame(info_frame)
        time_frame.grid(row=2, column=0, columnspan=5, sticky=tk.W, pady=5)

        ttk.Label(time_frame, text=f"开始时间: {start_time if start_time else '未设置'}").pack(side=tk.LEFT, padx=10)
        if end_time:
            ttk.Label(time_frame, text=f"结束时间: {end_time}").pack(side=tk.LEFT, padx=10)

        # 状态
        ttk.Label(info_frame, text=f"状态: {status}").grid(row=3, column=0, sticky=tk.W, pady=5)

        # 任务列表区域
        task_frame = ttk.LabelFrame(main_frame, text="任务列表", padding=10)
        task_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 创建树形视图显示任务
        columns = ('created_at', 'task_id', 'description')
        self.task_tree = ttk.Treeview(task_frame, columns=columns, selectmode='browse', show='tree headings')
        self.task_tree.heading('#0', text='任务名称')
        self.task_tree.heading('created_at', text='创建时间')

        # 隐藏额外的列
        self.task_tree.column('task_id', width=0, stretch=False)
        self.task_tree.column('description', width=0, stretch=False)

        self.task_tree.column('#0', width=300)
        self.task_tree.column('created_at', width=200)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(task_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=scrollbar.set)

        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定双击事件
        self.task_tree.bind('<Double-1>', self.on_task_double_click)

        # 保存当前阶段ID
        self.current_stage_id = stage_id
        # 加载任务数据
        self.load_tasks(stage_id)

    def delete_stage_from_detail(self, stage_id, stage_name):
        """从阶段详情页面删除阶段"""
        if not messagebox.askyesno("确认删除", f"确定要删除阶段 '{stage_name}' 吗？\n这将删除该阶段的所有任务！"):
            return

        try:
            # 首先获取项目ID，以便返回项目详情
            self.cursor.execute("SELECT project_id FROM stages WHERE id = ?", (stage_id,))
            result = self.cursor.fetchone()
            if not result:
                messagebox.showerror("错误", "找不到对应的项目")
                return

            project_id = result[0]

            # 获取项目名称
            self.cursor.execute("SELECT name FROM projects WHERE id = ?", (project_id,))
            project_result = self.cursor.fetchone()
            if not project_result:
                messagebox.showerror("错误", "找不到对应的项目名称")
                return

            project_name = project_result[0]

            # 删除阶段（由于外键约束，相关的任务会自动删除）
            self.cursor.execute("DELETE FROM stages WHERE id = ?", (stage_id,))

            # 更新项目更新时间
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute("UPDATE projects SET updated_at = ? WHERE id = ?",
                                (current_time, project_id))

            self.conn.commit()

            # 直接返回项目详情页面，不尝试刷新其他可能不存在的控件
            self.show_project_detail(project_id, project_name)

            messagebox.showinfo("成功", f"阶段 '{stage_name}' 已删除")

        except Exception as e:
            messagebox.showerror("错误", f"删除阶段时出错:\n{str(e)}")

    def save_stage_name(self):
        """保存阶段名称修改"""
        new_name = self.stage_name_var.get()
        if new_name:
            self.cursor.execute("UPDATE stages SET name = ? WHERE id = ?",
                                (new_name, self.current_stage_id))
            self.conn.commit()
            # 刷新阶段视图
            self.load_stages(self.current_project_id)

    def complete_stage(self):
        """结束阶段"""
        # 获取当前阶段状态
        self.cursor.execute("SELECT status FROM stages WHERE id = ?", (self.current_stage_id,))
        status = self.cursor.fetchone()[0]

        if status == '已完成':
            messagebox.showinfo("提示", "该阶段已经完成")
            return

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.cursor.execute("UPDATE stages SET end_time = ?, status = '已完成' WHERE id = ?",
                            (current_time, self.current_stage_id))

        # 更新项目更新时间
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.cursor.execute("UPDATE projects SET updated_at = ? WHERE id = ?",
                            (current_time, self.current_project_id))

        self.conn.commit()

        # 刷新视图 - 先检查stage_canvas是否存在
        if hasattr(self, 'stage_canvas') and self.stage_canvas.winfo_exists():
            self.load_stages(self.current_project_id)

        # 刷新阶段详情视图
        self.show_stage_detail(self.current_stage_id, self.stage_name_var.get())

        # 刷新项目列表中的进度（新增）
        self.load_projects()

    def add_task(self, stage_id):
        """添加新任务"""
        name = simpledialog.askstring("新增任务", "请输入任务名称:")
        if name:
            self.cursor.execute("INSERT INTO tasks (stage_id, name) VALUES (?, ?)",
                                (stage_id, name))
            self.conn.commit()
            self.load_tasks(stage_id)

    def load_tasks(self, stage_id):
        """加载任务数据"""
        # 检查task_tree是否存在并且有效
        if not hasattr(self, 'task_tree'):
            return

        # 检查控件是否仍然存在
        try:
            if not self.task_tree.winfo_exists():
                return
        except:
            return  # 如果检查失败，说明控件已不存在

        # 清空现有数据
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)

        # 从数据库获取任务数据
        self.cursor.execute('''
            SELECT id, name, description, created_at 
            FROM tasks 
            WHERE stage_id = ? 
            ORDER BY created_at
        ''', (stage_id,))

        tasks = self.cursor.fetchall()

        for task in tasks:
            task_id, name, description, created_at = task
            item_id = self.task_tree.insert('', 'end', text=name,
                                            values=(created_at, task_id, description))

    def on_task_double_click(self, event):
        """双击任务事件"""
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择一个任务")
            return

        item = selected_items[0]
        values = self.task_tree.item(item, 'values')
        if not values or len(values) < 2:
            messagebox.showerror("错误", "任务数据不完整")
            return

        task_id = values[1]
        task_name = self.task_tree.item(item, 'text')
        description = values[2] if len(values) > 2 else ""

        # 获取阶段名称（简化处理，实际应该存储在实例变量中）
        self.cursor.execute("SELECT name FROM stages WHERE id = ?", (self.current_stage_id,))
        stage_name = self.cursor.fetchone()[0]

        # 保存到导航历史
        self.navigation_history.append(('stage_detail', (self.current_stage_id, stage_name)))

        # 显示任务详情
        self.navigate_to('task_detail', task_id, task_name, description)

    def show_task_detail(self, task_id, task_name, description):
        """显示任务详情面板"""
        # 清除现有控件
        for widget in self.root.winfo_children():
            widget.destroy()

        self.current_view = ('task_detail', task_id)

        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        back_btn = ttk.Button(toolbar, text="← 返回阶段详情", command=self.navigate_back)
        back_btn.pack(side=tk.LEFT, padx=2)

        ttk.Label(toolbar, text=f"任务: {task_name}", font=('Arial', 12, 'bold')).pack(side=tk.LEFT, padx=10)

        # 主内容区域
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 任务信息区域
        info_frame = ttk.LabelFrame(main_frame, text="任务信息", padding=10)
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="任务名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.task_name_var = tk.StringVar(value=task_name)
        name_entry = ttk.Entry(info_frame, textvariable=self.task_name_var, width=40)
        name_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # 按钮放在第二行
        button_frame = ttk.Frame(info_frame)
        button_frame.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=5)

        save_btn = ttk.Button(button_frame, text="保存", command=self.save_task)
        save_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(button_frame, text="删除任务", command=lambda: self.delete_task(task_id))
        delete_btn.pack(side=tk.LEFT, padx=5)

        # 任务描述区域
        desc_frame = ttk.LabelFrame(main_frame, text="任务描述", padding=10)
        desc_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.desc_text = tk.Text(desc_frame, wrap=tk.WORD, height=10)
        self.desc_text.pack(fill=tk.BOTH, expand=True, pady=5)
        if description:
            self.desc_text.insert('1.0', description)

        # 绑定文本变化事件
        self.desc_text.bind('<KeyRelease>', self.on_description_changed)

        # 保存当前任务ID和原始描述
        self.current_task_id = task_id
        self.original_task_name = task_name
        self.original_description = description
        self.changes_made = False

        # 保存当前阶段ID，因为我们需要它来返回阶段视图
        self.cursor.execute("SELECT stage_id FROM tasks WHERE id = ?", (task_id,))
        self.current_stage_id = self.cursor.fetchone()[0]

    def on_description_changed(self, event):
        """任务描述变化时调用"""
        current_description = self.desc_text.get('1.0', tk.END).strip()
        self.changes_made = self.changes_made or (current_description != self.original_description)

    def save_task(self):
        """保存任务修改"""
        new_name = self.task_name_var.get()
        new_description = self.desc_text.get('1.0', tk.END).strip()

        if new_name and (new_name != self.original_task_name or new_description != self.original_description):
            self.cursor.execute("UPDATE tasks SET name = ?, description = ? WHERE id = ?",
                                (new_name, new_description, self.current_task_id))
            self.conn.commit()

            # 更新原始值
            self.original_task_name = new_name
            self.original_description = new_description
            self.changes_made = False

            # 刷新任务列表
            if hasattr(self, 'task_tree') and self.task_tree.winfo_exists():
                self.load_tasks(self.current_stage_id)

    def delete_task(self, task_id):
        """删除任务"""
        if messagebox.askyesno("确认删除", "确定要删除这个任务吗？"):
            self.cursor.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
            self.conn.commit()
            # 返回任务列表
            self.show_stage_detail(self.current_stage_id, self.stage_name_var.get())


if __name__ == "__main__":
    root = tk.Tk()
    app = ProjectManager(root)
    root.mainloop()