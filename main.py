import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import os
import pandas as pd
from pypinyin import pinyin, Style
import configparser


class TrophyManager:
    def __init__(self, root):
        self.root = root
        self.root.title("战利品管理器")
        self.root.geometry("900x600")

        # 设置字体
        self.font_style = ("微软雅黑", 11)
        self.header_font_style = ("微软雅黑", 11, "bold")

        # 排序选项和当前排序方式
        self.sort_options = {
            "物种升序": ("species", True),
            "物种降序": ("species", False),
            "等级升序": ("grade", True),
            "等级降序": ("grade", False),
            "评分升序": ("score", True),
            "评分降序": ("score", False)
        }

        self.current_sort = "物种升序"

        # 设置图标
        self.set_window_icon()

        # 默认设置
        self.settings_file = "settings.ini"
        self.settings_section = "TrophyManager"

        # 加载设置
        self.load_settings()

        # 创建GUI
        self.create_widgets()

        # 加载数据
        self.load_data()

    def set_window_icon(self):
        """设置窗口图标"""
        icon_paths = 'icon/COTW.ico'  # resources子目录
        if os.path.exists(icon_paths):
            self.root.iconbitmap(icon_paths)

    def load_settings(self):
        """从INI文件加载设置"""
        config = configparser.ConfigParser()
        # 设置默认值
        config[self.settings_section] = {
            "csv_path": "trophy.csv"
        }

        try:
            if os.path.exists(self.settings_file):
                config.read(self.settings_file, encoding="utf-8")

            self.settings = {
                "csv_path": config.get(self.settings_section, "csv_path", fallback="trophies.csv")
            }
        except Exception as e:
            print(f"加载设置失败: {e}")
            self.settings = {"csv_path": "trophies.csv"}

    def save_settings(self):
        """保存设置到INI文件"""
        config = configparser.ConfigParser()
        config[self.settings_section] = {
            "csv_path": self.settings["csv_path"]
        }

        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                config.write(f)
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {e}")

    def create_widgets(self):
        """创建主界面控件"""
        # 顶部工具栏
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        # 刷新按钮
        refresh_btn = tk.Button(toolbar, text="刷新", command=self.load_data)
        refresh_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # 添加按钮
        add_btn = tk.Button(toolbar, text="添加", command=self.show_add_dialog)
        add_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # 删除按钮
        delete_btn = tk.Button(toolbar, text="删除", command=self.delete_selected)
        delete_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # 设置按钮
        settings_btn = tk.Button(toolbar, text="设置", command=self.show_settings_dialog)
        settings_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # 搜索框
        search_frame = tk.Frame(toolbar)
        search_frame.pack(side=tk.LEFT, padx=5)

        tk.Label(search_frame, text="搜索物种:").pack(side=tk.LEFT)
        self.search_entry = tk.Entry(search_frame, width=20)
        self.search_entry.pack(side=tk.LEFT, padx=2)

        # 绑定Enter键事件
        self.search_entry.bind("<Return>", lambda event: self.search_data())

        search_btn = tk.Button(search_frame, text="搜索", command=self.search_data)
        search_btn.pack(side=tk.LEFT, padx=2)

        # 在工具栏右侧添加排序控件
        sort_frame = tk.Frame(toolbar)
        sort_frame.pack(side=tk.RIGHT, padx=5)

        tk.Label(sort_frame, text="排序方式:").pack(side=tk.LEFT)

        # 排序方式下拉框
        self.sort_var = tk.StringVar(value=self.current_sort)
        self.sort_menu = ttk.Combobox(
            sort_frame,
            textvariable=self.sort_var,
            values=list(self.sort_options.keys()),
            state="readonly",
            width=16
        )
        self.sort_menu.pack(side=tk.LEFT, padx=2)
        self.sort_menu.bind("<<ComboboxSelected>>", self.change_sort_method)

        # 数据表格
        self.create_table()

    def create_table(self):
        """创建数据表格"""
        # 表格框架
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 滚动条
        scroll_y = tk.Scrollbar(table_frame)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        scroll_x = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 表格
        self.table = ttk.Treeview(
            table_frame,
            columns=("species", "color", "grade", "score", "id"),
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        # 配置滚动条
        scroll_y.config(command=self.table.yview)
        scroll_x.config(command=self.table.xview)

        # 定义列
        self.table.column("#0", width=0, stretch=tk.NO)
        self.table.column("species", width=200, anchor=tk.CENTER)
        self.table.column("color", width=150, anchor=tk.CENTER)
        self.table.column("grade", width=100, anchor=tk.CENTER)
        self.table.column("score", width=100, anchor=tk.CENTER)
        self.table.column("id", width=80, anchor=tk.CENTER)

        # 定义表头
        self.table.heading("species", text="物种")
        self.table.heading("color", text="毛色")
        self.table.heading("grade", text="等级")
        self.table.heading("score", text="评分")
        self.table.heading("id", text="ID")

        # 设置表头样式
        style = ttk.Style()
        style.configure("Treeview.Heading", font=self.header_font_style)
        style.configure("Treeview", font=self.font_style, rowheight=25)

        # 绑定双击事件
        self.table.bind("<Double-1>", self.on_row_double_click)

        # 绑定双击事件
        self.table.bind("<Double-1>", self.on_row_double_click)

        # 创建右键菜单
        self.create_context_menu()

        # 绑定右键点击事件
        self.table.bind("<Button-3>", self.show_context_menu)

        self.table.pack(fill=tk.BOTH, expand=True)

    def change_sort_method(self, event=None):
        """更改排序方法"""
        self.current_sort = self.sort_var.get()
        self.load_data()

    def load_data(self):
        """从CSV文件加载数据并显示在表格中"""
        try:
            # 清空表格
            for item in self.table.get_children():
                self.table.delete(item)

            # 检查文件是否存在
            if not os.path.exists(self.settings["csv_path"]):
                # 如果文件不存在，创建一个空的
                with open(self.settings["csv_path"], "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["species", "color", "grade", "score", "id"])
                return

            # 读取CSV文件
            df = pd.read_csv(self.settings["csv_path"], encoding="utf-8-sig")

            # 获取当前排序方式
            sort_key, ascending = self.sort_options[self.current_sort]

            # 添加拼音列用于排序
            df['pinyin'] = df['species'].apply(
                lambda x: ''.join([i[0] for i in pinyin(x, style=Style.FIRST_LETTER)]))

            # 定义等级排序权重
            grade_order = {
                "珍禽异兽": 5,
                "钻石": 4,
                "黄金": 3,
                "白银": 2,
                "青铜": 1
            }
            df['grade_weight'] = df['grade'].map(grade_order)

            # 根据排序键选择排序列
            if sort_key == "species":
                sort_col = 'pinyin'
            elif sort_key == "grade":
                sort_col = 'grade_weight'
            else:
                sort_col = sort_key

            # 主排序
            df_sorted = df.sort_values(
                sort_col,
                ascending=ascending,
                kind='mergesort'  # 保持排序稳定性
            )

            # # 次级排序（等级降序、评分降序、ID升序）
            # df_sorted = df_sorted.sort_values(
            #     ['grade_weight', 'score', 'id'],
            #     ascending=[False, False, True],
            #     kind='mergesort'
            # )

            # 添加数据到表格
            for _, row in df_sorted.iterrows():
                self.add_row_to_table(row)

        except Exception as e:
            messagebox.showerror("错误", f"加载数据失败: {e}")

    def add_row_to_table(self, row):
        """添加一行数据到表格"""
        grade = row["grade"]
        # 根据等级设置颜色
        color = {
            "青铜": "#CD7F32",
            "白银": "#C0C0C0",
            "黄金": "#FFD700",
            "钻石": "#B9F2FF",
            "珍禽异兽": "#800080"
        }.get(grade, "black")

        # 格式化评分为两位小数
        score = "{:.2f}".format(float(row["score"]))

        self.table.insert("", tk.END, values=(
            row["species"],
            row["color"],
            grade,
            score,
            row["id"]
        ), tags=(color,))

        # 设置行颜色
        self.table.tag_configure("#CD7F32", foreground="#CD7F32")
        self.table.tag_configure("#C0C0C0", foreground="#C0C0C0")
        self.table.tag_configure("#FFD700", foreground="#FFD700")
        self.table.tag_configure("#B9F2FF", foreground="#B9F2FF")
        self.table.tag_configure("#800080", foreground="#800080")

    def search_data(self):
        """搜索数据"""
        keyword = self.search_entry.get().strip()
        if not keyword:
            self.load_data()
            return

        try:
            # 清空表格
            for item in self.table.get_children():
                self.table.delete(item)

            # 读取CSV文件
            df = pd.read_csv(self.settings["csv_path"], encoding="utf-8-sig")

            # 模糊搜索
            mask = df["species"].str.contains(keyword, case=False, na=False)
            result = df.loc[mask].copy()

            # 获取当前排序方式
            sort_key, ascending = self.sort_options.get(
                self.current_sort, ("species", True)  # 默认按物种拼音升序
            )

            # 添加拼音列用于排序
            result['pinyin'] = result['species'].apply(
                lambda x: ''.join([i[0] for i in pinyin(x, style=Style.FIRST_LETTER)]))

            # 定义等级排序权重
            grade_order = {
                "珍禽异兽": 5,
                "钻石": 4,
                "黄金": 3,
                "白银": 2,
                "青铜": 1
            }
            result['grade_weight'] = result['grade'].map(grade_order)

            # 根据排序键选择排序列
            if sort_key == "species":
                sort_col = 'pinyin'
            elif sort_key == "grade":
                sort_col = 'grade_weight'
            else:
                sort_col = sort_key

            # 主排序
            result_sorted = result.sort_values(
                sort_col,
                ascending=ascending,
                kind='mergesort'
            )

            # 次级排序（等级降序、评分降序、ID升序）
            result_sorted = result_sorted.sort_values(
                ['grade_weight', 'score', 'id'],
                ascending=[False, False, True],
                kind='mergesort'
            )

            # 添加数据到表格
            for _, row in result_sorted.iterrows():
                self.add_row_to_table(row)

        except Exception as e:
            messagebox.showerror("错误", f"搜索数据失败: {e}")

    def get_next_id(self):
        """获取下一个可用的ID（当前最大ID + 1）"""
        try:
            if not os.path.exists(self.settings["csv_path"]):
                return 1

            max_id = 0
            with open(self.settings["csv_path"], "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        current_id = int(row["id"])
                        if current_id > max_id:
                            max_id = current_id
                    except (ValueError, KeyError):
                        continue

            return max_id + 1
        except Exception as e:
            messagebox.showerror("错误", f"获取ID失败: {e}")
            return 1

    def show_add_dialog(self):
        """显示添加战利品的对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加战利品")
        dialog.transient(self.root)
        dialog.grab_set()

        # 设置对话框尺寸并居中
        dialog_width = 300
        dialog_height = 200
        self.center_window(dialog, dialog_width, dialog_height)

        # 物种
        tk.Label(dialog, text="物种:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        species_entry = tk.Entry(dialog)
        species_entry.grid(row=0, column=1, padx=5, pady=5)

        # 毛色
        tk.Label(dialog, text="毛色:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        color_entry = tk.Entry(dialog)
        color_entry.grid(row=1, column=1, padx=5, pady=5)

        # 等级
        tk.Label(dialog, text="等级:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        grade_var = tk.StringVar(dialog)
        grade_var.set("青铜")
        grade_menu = tk.OptionMenu(dialog, grade_var, "青铜", "白银", "黄金", "钻石", "珍禽异兽")
        grade_menu.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        # 评分
        tk.Label(dialog, text="评分:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        score_entry = tk.Entry(dialog)
        score_entry.grid(row=3, column=1, padx=5, pady=5)

        # 按钮
        button_frame = tk.Frame(dialog)
        button_frame.grid(row=4, column=0, columnspan=2, pady=5)

        def add_trophy():
            """添加战利品"""
            species = species_entry.get().strip()
            color = color_entry.get().strip()
            grade = grade_var.get()
            score = score_entry.get().strip()

            # 验证数据
            if not species or not color or not score:
                messagebox.showerror("错误", "所有字段都必须填写")
                return

            try:
                score = float(score)
                if score < 0:
                    raise ValueError("评分不能为负数")
            except ValueError:
                messagebox.showerror("错误", "评分必须是数字且不小于0")
                return

            # 获取下一个ID
            next_id = self.get_next_id()

            # 格式化评分为两位小数
            score = "{:.2f}".format(score)

            # 写入CSV文件
            try:
                file_exists = os.path.exists(self.settings["csv_path"]) and os.path.getsize(
                    self.settings["csv_path"]) > 0
                with open(self.settings["csv_path"], "a", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    if not file_exists:
                        writer.writerow(["species", "color", "grade", "score", "id"])
                    writer.writerow([species, color, grade, score, next_id])

                # 刷新表格
                self.load_data()
                dialog.destroy()

            except Exception as e:
                messagebox.showerror("错误", f"保存数据失败: {e}")

        tk.Button(button_frame, text="确认", command=add_trophy).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def delete_selected(self):
        """删除选中的一条或多条记录"""
        selected_items = self.table.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的战利品")
            return

        # 获取选中的ID列表
        selected_ids = [self.table.item(item, "values")[4] for item in selected_items]

        # 确认删除
        confirm = messagebox.askyesno(
            "确认删除",
            f"确定要删除选中的 {len(selected_ids)} 个战利品吗？"
        )
        if not confirm:
            return

        try:
            # 读取所有数据
            with open(self.settings["csv_path"], "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                fieldnames = [field.strip('\ufeff') for field in reader.fieldnames]
                # 保留未选中的记录
                rows = [row for row in reader if row["id"] not in selected_ids]

            # 写回文件
            with open(self.settings["csv_path"], "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            # 刷新表格
            self.load_data()
            messagebox.showinfo("成功", f"已删除 {len(selected_ids)} 个战利品")

        except Exception as e:
            messagebox.showerror("错误", f"删除战利品失败: {e}")

    def on_row_double_click(self, event):
        """双击行事件处理"""
        item = self.table.selection()[0]
        values = self.table.item(item, "values")

        # 显示修改对话框
        self.show_edit_dialog(values)

    def create_context_menu(self):
        """创建右键上下文菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="修改", command=self.edit_selected_row)
        self.context_menu.add_command(label="删除", command=self.delete_selected_row)

    def show_context_menu(self, event):
        """显示右键上下文菜单"""
        item = self.table.identify_row(event.y)
        if item:
            self.table.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def edit_selected_row(self):
        """编辑选中的行"""
        selected_items = self.table.selection()
        if selected_items:
            item = selected_items[0]
            values = self.table.item(item, "values")
            self.show_edit_dialog(values)

    def delete_selected_row(self):
        """删除选中的行"""
        selected_items = self.table.selection()
        if not selected_items:
            return

        item = selected_items[0]
        values = self.table.item(item, "values")
        trophy_name = values[0]
        trophy_id =values[4]

        # 确认删除
        if not messagebox.askyesno("确认删除", f"确定要删除 {trophy_name} 的战利品吗？"):
            return

        try:
            # 读取所有数据
            with open(self.settings["csv_path"], "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                fieldnames = [field.strip('\ufeff') for field in reader.fieldnames]
                rows = [row for row in reader if row["id"] != trophy_id]

            # 写回文件
            with open(self.settings["csv_path"], "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)

            # 刷新表格
            self.load_data()
            messagebox.showinfo("成功", "战利品已删除")

        except Exception as e:
            messagebox.showerror("错误", f"删除数据失败: {e}")

    def show_edit_dialog(self, values):
        """显示修改战利品的对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("修改战利品")
        dialog.transient(self.root)
        dialog.grab_set()

        # 设置对话框尺寸并居中
        dialog_width = 300
        dialog_height = 200
        self.center_window(dialog, dialog_width, dialog_height)

        # 显示物种和毛色（不可编辑）
        tk.Label(dialog, text="物种:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        tk.Label(dialog, text=values[0]).grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        tk.Label(dialog, text="毛色:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        tk.Label(dialog, text=values[1]).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        # # 显示ID（不可编辑）
        # tk.Label(dialog, text="ID:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        # tk.Label(dialog, text=values[4]).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        # 等级（可修改）
        tk.Label(dialog, text="等级:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        grade_var = tk.StringVar(dialog)
        grade_var.set(values[2])
        grade_menu = tk.OptionMenu(dialog, grade_var, "青铜", "白银", "黄金", "钻石", "珍禽异兽")
        grade_menu.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)

        # 评分（可修改）
        tk.Label(dialog, text="评分:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.E)
        score_entry = tk.Entry(dialog)
        score_entry.insert(0, values[3])
        score_entry.grid(row=4, column=1, padx=5, pady=5)

        # 按钮
        button_frame = tk.Frame(dialog)
        button_frame.grid(row=5, column=0, columnspan=2, pady=5)

        def update_trophy():
            """更新战利品信息"""
            species = values[0]
            color = values[1]
            trophy_id = values[4]
            grade = grade_var.get()
            score = score_entry.get().strip()

            # 验证数据
            if not score:
                messagebox.showerror("错误", "评分必须填写")
                return

            try:
                score = float(score)
                if score < 0:
                    raise ValueError("评分不能为负数")
            except ValueError:
                messagebox.showerror("错误", "评分必须是数字且不小于0")
                return

            # 格式化评分为两位小数
            score = "{:.2f}".format(score)

            # 读取所有数据
            try:
                with open(self.settings["csv_path"], "r", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)
                    rows = list(reader)

                # 更新匹配的行
                updated = False
                for row in rows:
                    if row["id"] == trophy_id:
                        row["grade"] = grade
                        row["score"] = score
                        updated = True
                        break

                if not updated:
                    messagebox.showerror("错误", "未找到匹配的战利品")
                    return

                # 写回文件
                with open(self.settings["csv_path"], "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=["species", "color", "grade", "score", "id"])
                    writer.writeheader()
                    writer.writerows(rows)

                # 刷新表格
                self.load_data()
                dialog.destroy()

            except Exception as e:
                messagebox.showerror("错误", f"更新数据失败: {e}")

        tk.Button(button_frame, text="确认", command=update_trophy).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def show_settings_dialog(self):
        """显示设置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("设置")
        dialog.transient(self.root)
        dialog.grab_set()

        # 设置对话框尺寸并居中
        dialog_width = 400
        dialog_height = 100
        self.center_window(dialog, dialog_width, dialog_height)

        # CSV文件路径
        tk.Label(dialog, text="CSV文件路径:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        path_entry = tk.Entry(dialog, width=20)
        path_entry.insert(0, self.settings["csv_path"])
        path_entry.grid(row=0, column=1, padx=5, pady=5)

        # 浏览按钮
        def browse_path():
            path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV文件", "*.csv")],
                initialfile="trophies.csv"
            )
            if path:
                path_entry.delete(0, tk.END)
                path_entry.insert(0, path)

        tk.Button(dialog, text="浏览", command=browse_path).grid(row=0, column=2, padx=5, pady=5)

        # 按钮
        button_frame = tk.Frame(dialog)
        button_frame.grid(row=1, column=0, columnspan=3, pady=5)

        def save_settings():
            """保存设置"""
            path = path_entry.get().strip()

            if not path:
                messagebox.showerror("错误", "CSV文件路径不能为空")
                return

            # 检查路径是否有效
            try:
                dirname = os.path.dirname(path)
                if dirname and not os.path.exists(dirname):
                    os.makedirs(dirname)
            except Exception as e:
                messagebox.showerror("错误", f"创建目录失败: {e}")
                return

            # 更新设置
            self.settings["csv_path"] = path
            self.save_settings()

            # 刷新数据
            self.load_data()
            dialog.destroy()

        tk.Button(button_frame, text="确认", command=save_settings).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def center_window(self, window, width=None, height=None):
        """将窗口居中显示在父窗口中心"""
        window.update_idletasks()  # 确保窗口尺寸已更新

        # 获取父窗口位置和尺寸
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # 计算居中位置
        if width and height:
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
            window.geometry(f"{width}x{height}+{x}+{y}")
        else:
            # 使用窗口自然尺寸
            width = window.winfo_width()
            height = window.winfo_height()
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2
            window.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TrophyManager(root)
    root.mainloop()