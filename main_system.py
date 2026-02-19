import sqlite3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

class DebateTeamManagementSystem:
    def __init__(self, db_name="debate_team.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(self.db_name)
        self.cursor = self.conn.cursor()
        self.initialize_database()
    
    def initialize_database(self):
        """初始化数据库表"""
        # 队员信息表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                position TEXT,
                join_date DATE,
                experience_level TEXT
            )
        ''')
        
        # 比赛记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                opponent TEXT NOT NULL,
                tournament TEXT,
                result TEXT CHECK(result IN ('Win', 'Loss', 'Draw')),
                score TEXT
            )
        ''')
        
        # 出战记录表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS match_participation (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER,
                match_id INTEGER,
                role TEXT,
                performance_score REAL,
                FOREIGN KEY (member_id) REFERENCES members(id),
                FOREIGN KEY (match_id) REFERENCES matches(id)
            )
        ''')
        
        # 排班表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedule (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                time_slot TEXT,
                activity TEXT,
                assigned_member_id INTEGER,
                FOREIGN KEY (assigned_member_id) REFERENCES members(id)
            )
        ''')
        
        self.conn.commit()
    
    def add_member(self, name, position, join_date, experience_level):
        """添加队员"""
        self.cursor.execute('''
            INSERT INTO members (name, position, join_date, experience_level)
            VALUES (?, ?, ?, ?)
        ''', (name, position, join_date, experience_level))
        self.conn.commit()
        return self.cursor.lastrowid
    
    def add_match(self, date, opponent, tournament, result, score):
        """添加比赛记录"""
        self.cursor.execute('''
            INSERT INTO matches (date, opponent, tournament, result, score)
            VALUES (?, ?, ?, ?, ?)
        ''', (date, opponent, tournament, result, score))
        self.conn.commit()
        return self.cursor.lastrowid
    
    def record_participation(self, member_id, match_id, role, performance_score):
        """记录队员出战情况"""
        self.cursor.execute('''
            INSERT INTO match_participation (member_id, match_id, role, performance_score)
            VALUES (?, ?, ?, ?)
        ''', (member_id, match_id, role, performance_score))
        self.conn.commit()
    
    def create_schedule(self, date, time_slot, activity, assigned_member_id):
        """创建排班表"""
        self.cursor.execute('''
            INSERT INTO schedule (date, time_slot, activity, assigned_member_id)
            VALUES (?, ?, ?, ?)
        ''', (date, time_slot, activity, assigned_member_id))
        self.conn.commit()
    
    def get_all_members(self):
        """获取所有队员信息"""
        query = "SELECT * FROM members"
        df = pd.read_sql_query(query, self.conn)
        return df
    
    def get_all_matches(self):
        """获取所有比赛信息"""
        query = "SELECT * FROM matches"
        df = pd.read_sql_query(query, self.conn)
        return df
    
    def get_member_stats(self):
        """获取队员统计数据"""
        query = '''
            SELECT 
                m.id,
                m.name,
                m.position,
                m.experience_level,
                COUNT(mp.id) as matches_played,
                AVG(mp.performance_score) as avg_performance,
                SUM(CASE WHEN ma.result = 'Win' THEN 1 ELSE 0 END) as wins
            FROM members m
            LEFT JOIN match_participation mp ON m.id = mp.member_id
            LEFT JOIN matches ma ON mp.match_id = ma.id
            GROUP BY m.id
        '''
        df = pd.read_sql_query(query, self.conn)
        return df
    
    def get_schedule(self):
        """获取排班表"""
        query = '''
            SELECT 
                s.date,
                s.time_slot,
                s.activity,
                m.name as assigned_member
            FROM schedule s
            LEFT JOIN members m ON s.assigned_member_id = m.id
            ORDER BY s.date
        '''
        df = pd.read_sql_query(query, self.conn)
        return df
    
    def update_member(self, member_id, name, position, join_date, experience_level):
        """更新队员信息"""
        self.cursor.execute('''
            UPDATE members 
            SET name=?, position=?, join_date=?, experience_level=?
            WHERE id=?
        ''', (name, position, join_date, experience_level, member_id))
        self.conn.commit()
    
    def delete_member(self, member_id):
        """删除队员"""
        # 删除相关的出战记录
        self.cursor.execute('DELETE FROM match_participation WHERE member_id=?', (member_id,))
        # 删除相关的排班记录
        self.cursor.execute('DELETE FROM schedule WHERE assigned_member_id=?', (member_id,))
        # 删除队员
        self.cursor.execute('DELETE FROM members WHERE id=?', (member_id,))
        self.conn.commit()
    
    def generate_performance_report(self):
        """生成能力评估报告"""
        stats_df = self.get_member_stats()
        
        # 保存数据到CSV供R分析
        stats_df.to_csv('performance_data.csv', index=False)
        
        # 运行R脚本进行高级分析
        r_script = '''
        # R_analysis.R
        library(ggplot2)
        library(dplyr)
        
        # 读取数据
        data <- read.csv("performance_data.csv")
        
        # 清理数据
        data <- data[!is.na(data$avg_performance), ]
        
        # 绘制性能图
        p1 <- ggplot(data, aes(x=reorder(name, avg_performance), y=avg_performance)) +
          geom_bar(stat="identity", fill="steelblue") +
          coord_flip() +
          labs(title="队员平均表现评分", x="队员姓名", y="平均表现评分") +
          theme_minimal()
        
        # 绘制胜率图
        p2 <- ggplot(data, aes(x=reorder(name, matches_played), y=matches_played)) +
          geom_bar(stat="identity", fill="orange") +
          coord_flip() +
          labs(title="队员参赛次数", x="队员姓名", y="参赛次数") +
          theme_minimal()
        
        # 保存图表
        ggsave("performance_chart.png", plot=p1, width=10, height=6)
        ggsave("participation_chart.png", plot=p2, width=10, height=6)
        
        # 输出摘要
        if(nrow(data) > 0) {
            cat("=== 队员能力评估摘要 ===\n")
            cat("平均表现最高:", data$name[which.max(data$avg_performance)], "\n")
            cat("参赛最多:", data$name[which.max(data$matches_played)], "\n")
            cat("胜场最多:", data$name[which.max(data$wins)], "\n")
        }
        '''
        
        # 将R脚本写入文件
        with open('R_analysis.R', 'w') as f:
            f.write(r_script)
        
        # 尝试运行R脚本
        try:
            result = subprocess.run(['Rscript', 'R_analysis.R'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                print("R分析完成，图表已生成")
            else:
                print("R脚本执行失败，错误:", result.stderr)
        except FileNotFoundError:
            print("R未安装或未添加到PATH中，跳过R分析")
        
        return stats_df
    
    def export_data(self, format_type='csv'):
        """导出数据"""
        tables = ['members', 'matches', 'match_participation', 'schedule']
        for table in tables:
            df = pd.read_sql_query(f'SELECT * FROM {table}', self.conn)
            if format_type.lower() == 'csv':
                df.to_csv(f'{table}_export.csv', index=False)
            elif format_type.lower() == 'excel':
                df.to_excel(f'{table}_export.xlsx', index=False)
        messagebox.showinfo("导出成功", f"数据已导出为{format_type.upper()}格式")
    
    def close_connection(self):
        """关闭数据库连接"""
        self.conn.close()

class DebateTeamApp:
    def __init__(self, root):
        self.root = root
        self.root.title("辩论队电子信息管理系统")
        self.root.geometry("1000x700")
        
        self.system = DebateTeamManagementSystem()
        
        self.setup_ui()
        self.load_initial_data()
    
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 队员管理选项卡
        self.members_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.members_frame, text="队员管理")
        self.setup_members_tab()
        
        # 比赛管理选项卡
        self.matches_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.matches_frame, text="比赛管理")
        self.setup_matches_tab()
        
        # 出战统计选项卡
        self.participation_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.participation_frame, text="出战统计")
        self.setup_participation_tab()
        
        # 排班表选项卡
        self.schedule_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.schedule_frame, text="排班表")
        self.setup_schedule_tab()
        
        # 数据分析选项卡
        self.analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.analysis_frame, text="数据分析")
        self.setup_analysis_tab()
        
        # 导航按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="导出数据", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新数据", command=self.refresh_all_tabs).pack(side=tk.LEFT, padx=5)
    
    def setup_members_tab(self):
        # 输入区域
        input_frame = ttk.LabelFrame(self.members_frame, text="添加/编辑队员")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 姓名
        ttk.Label(input_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.member_name_var = tk.StringVar()
        self.member_name_entry = ttk.Entry(input_frame, textvariable=self.member_name_var)
        self.member_name_entry.grid(row=0, column=1, padx=5, pady=2)
        
        # 位置
        ttk.Label(input_frame, text="位置:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.member_position_var = tk.StringVar()
        self.member_position_combo = ttk.Combobox(input_frame, textvariable=self.member_position_var, 
                                                  values=["一辩", "二辩", "三辩", "四辩", "自由辩"])
        self.member_position_combo.grid(row=0, column=3, padx=5, pady=2)
        
        # 加入日期
        ttk.Label(input_frame, text="加入日期:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.member_join_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.member_join_date_entry = ttk.Entry(input_frame, textvariable=self.member_join_date_var)
        self.member_join_date_entry.grid(row=1, column=1, padx=5, pady=2)
        
        # 经验等级
        ttk.Label(input_frame, text="经验等级:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        self.member_experience_var = tk.StringVar()
        self.member_experience_combo = ttk.Combobox(input_frame, textvariable=self.member_experience_var, 
                                                    values=["初级", "中级", "高级"])
        self.member_experience_combo.grid(row=1, column=3, padx=5, pady=2)
        
        # 操作按钮
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=5)
        
        self.add_member_btn = ttk.Button(btn_frame, text="添加队员", command=self.add_member)
        self.add_member_btn.pack(side=tk.LEFT, padx=5)
        
        self.update_member_btn = ttk.Button(btn_frame, text="更新队员", command=self.update_member)
        self.update_member_btn.pack(side=tk.LEFT, padx=5)
        self.update_member_btn.config(state=tk.DISABLED)
        
        self.delete_member_btn = ttk.Button(btn_frame, text="删除队员", command=self.delete_member)
        self.delete_member_btn.pack(side=tk.LEFT, padx=5)
        self.delete_member_btn.config(state=tk.DISABLED)
        
        # 重置按钮
        self.reset_member_btn = ttk.Button(btn_frame, text="重置", command=self.reset_member_form)
        self.reset_member_btn.pack(side=tk.LEFT, padx=5)
        
        # 表格区域
        table_frame = ttk.LabelFrame(self.members_frame, text="队员列表")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ('ID', '姓名', '位置', '加入日期', '经验等级')
        self.members_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.members_tree.heading(col, text=col)
            self.members_tree.column(col, width=120)
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.members_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.members_tree.xview)
        self.members_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.members_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定选择事件
        self.members_tree.bind('<<TreeviewSelect>>', self.on_member_select)
    
    def setup_matches_tab(self):
        # 输入区域
        input_frame = ttk.LabelFrame(self.matches_frame, text="添加比赛")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 日期
        ttk.Label(input_frame, text="日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.match_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.match_date_entry = ttk.Entry(input_frame, textvariable=self.match_date_var)
        self.match_date_entry.grid(row=0, column=1, padx=5, pady=2)
        
        # 对手
        ttk.Label(input_frame, text="对手:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.match_opponent_var = tk.StringVar()
        self.match_opponent_entry = ttk.Entry(input_frame, textvariable=self.match_opponent_var)
        self.match_opponent_entry.grid(row=0, column=3, padx=5, pady=2)
        
        # 锦标赛
        ttk.Label(input_frame, text="锦标赛:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.match_tournament_var = tk.StringVar()
        self.match_tournament_entry = ttk.Entry(input_frame, textvariable=self.match_tournament_var)
        self.match_tournament_entry.grid(row=1, column=1, padx=5, pady=2)
        
        # 结果
        ttk.Label(input_frame, text="结果:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        self.match_result_var = tk.StringVar()
        self.match_result_combo = ttk.Combobox(input_frame, textvariable=self.match_result_var, 
                                               values=["Win", "Loss", "Draw"])
        self.match_result_combo.grid(row=1, column=3, padx=5, pady=2)
        
        # 分数
        ttk.Label(input_frame, text="分数:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.match_score_var = tk.StringVar()
        self.match_score_entry = ttk.Entry(input_frame, textvariable=self.match_score_var)
        self.match_score_entry.grid(row=2, column=1, padx=5, pady=2)
        
        # 操作按钮
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, pady=5)
        
        ttk.Button(btn_frame, text="添加比赛", command=self.add_match).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重置", command=self.reset_match_form).pack(side=tk.LEFT, padx=5)
        
        # 表格区域
        table_frame = ttk.LabelFrame(self.matches_frame, text="比赛列表")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ('ID', '日期', '对手', '锦标赛', '结果', '分数')
        self.matches_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.matches_tree.heading(col, text=col)
            self.matches_tree.column(col, width=120)
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.matches_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.matches_tree.xview)
        self.matches_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.matches_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_participation_tab(self):
        # 输入区域
        input_frame = ttk.LabelFrame(self.participation_frame, text="记录出战情况")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 队员选择
        ttk.Label(input_frame, text="队员:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.part_member_var = tk.StringVar()
        self.part_member_combo = ttk.Combobox(input_frame, textvariable=self.part_member_var)
        self.part_member_combo.grid(row=0, column=1, padx=5, pady=2)
        
        # 比赛选择
        ttk.Label(input_frame, text="比赛:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.part_match_var = tk.StringVar()
        self.part_match_combo = ttk.Combobox(input_frame, textvariable=self.part_match_var)
        self.part_match_combo.grid(row=0, column=3, padx=5, pady=2)
        
        # 角色
        ttk.Label(input_frame, text="角色:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.part_role_var = tk.StringVar()
        self.part_role_combo = ttk.Combobox(input_frame, textvariable=self.part_role_var, 
                                            values=["一辩", "二辩", "三辩", "四辩", "自由辩"])
        self.part_role_combo.grid(row=1, column=1, padx=5, pady=2)
        
        # 表现评分
        ttk.Label(input_frame, text="表现评分:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        self.part_score_var = tk.DoubleVar()
        self.part_score_scale = ttk.Scale(input_frame, from_=0, to=10, variable=self.part_score_var, 
                                          orient=tk.HORIZONTAL)
        self.part_score_scale.grid(row=1, column=3, padx=5, pady=2, sticky=tk.EW)
        
        # 操作按钮
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=5)
        
        ttk.Button(btn_frame, text="记录出战", command=self.record_participation).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重置", command=self.reset_participation_form).pack(side=tk.LEFT, padx=5)
        
        # 表格区域
        table_frame = ttk.LabelFrame(self.participation_frame, text="出战记录")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ('ID', '队员', '比赛', '角色', '表现评分')
        self.participation_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.participation_tree.heading(col, text=col)
            self.participation_tree.column(col, width=120)
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.participation_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.participation_tree.xview)
        self.participation_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.participation_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_schedule_tab(self):
        # 输入区域
        input_frame = ttk.LabelFrame(self.schedule_frame, text="创建排班")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 日期
        ttk.Label(input_frame, text="日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.schedule_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.schedule_date_entry = ttk.Entry(input_frame, textvariable=self.schedule_date_var)
        self.schedule_date_entry.grid(row=0, column=1, padx=5, pady=2)
        
        # 时间段
        ttk.Label(input_frame, text="时间段:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        self.schedule_time_var = tk.StringVar()
        self.schedule_time_combo = ttk.Combobox(input_frame, textvariable=self.schedule_time_var, 
                                                values=["上午", "下午", "晚上", "全天"])
        self.schedule_time_combo.grid(row=0, column=3, padx=5, pady=2)
        
        # 活动
        ttk.Label(input_frame, text="活动:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.schedule_activity_var = tk.StringVar()
        self.schedule_activity_entry = ttk.Entry(input_frame, textvariable=self.schedule_activity_var)
        self.schedule_activity_entry.grid(row=1, column=1, padx=5, pady=2)
        
        # 负责人
        ttk.Label(input_frame, text="负责人:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=2)
        self.schedule_member_var = tk.StringVar()
        self.schedule_member_combo = ttk.Combobox(input_frame, textvariable=self.schedule_member_var)
        self.schedule_member_combo.grid(row=1, column=3, padx=5, pady=2)
        
        # 操作按钮
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=5)
        
        ttk.Button(btn_frame, text="创建排班", command=self.create_schedule).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重置", command=self.reset_schedule_form).pack(side=tk.LEFT, padx=5)
        
        # 表格区域
        table_frame = ttk.LabelFrame(self.schedule_frame, text="排班表")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建树形视图
        columns = ('ID', '日期', '时间段', '活动', '负责人')
        self.schedule_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.schedule_tree.heading(col, text=col)
            self.schedule_tree.column(col, width=120)
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.schedule_tree.yview)
        tree_scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.schedule_tree.xview)
        self.schedule_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.schedule_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_analysis_tab(self):
        # 分析按钮
        btn_frame = ttk.Frame(self.analysis_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="生成能力评估报告", command=self.generate_performance_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成比赛统计图", command=self.generate_match_statistics).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成排班表", command=self.generate_schedule_report).pack(side=tk.LEFT, padx=5)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(self.analysis_frame, text="分析结果")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.analysis_text = tk.Text(result_frame, wrap=tk.WORD)
        text_scroll = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.analysis_text.yview)
        self.analysis_text.configure(yscrollcommand=text_scroll.set)
        
        self.analysis_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    def load_initial_data(self):
        # 加载初始数据到所有表格
        self.refresh_members_tab()
        self.refresh_matches_tab()
        self.refresh_participation_tab()
        self.refresh_schedule_tab()
    
    def refresh_members_tab(self):
        # 清空现有数据
        for item in self.members_tree.get_children():
            self.members_tree.delete(item)
        
        # 获取数据
        df = self.system.get_all_members()
        
        # 插入数据
        for _, row in df.iterrows():
            self.members_tree.insert('', tk.END, values=list(row))
    
    def refresh_matches_tab(self):
        # 清空现有数据
        for item in self.matches_tree.get_children():
            self.matches_tree.delete(item)
        
        # 获取数据
        df = self.system.get_all_matches()
        
        # 插入数据
        for _, row in df.iterrows():
            self.matches_tree.insert('', tk.END, values=list(row))
    
    def refresh_participation_tab(self):
        # 清空现有数据
        for item in self.participation_tree.get_children():
            self.participation_tree.delete(item)
        
        # 获取数据
        query = '''
            SELECT 
                mp.id,
                m.name as member_name,
                ma.opponent || ' (' || ma.date || ')' as match_info,
                mp.role,
                mp.performance_score
            FROM match_participation mp
            JOIN members m ON mp.member_id = m.id
            JOIN matches ma ON mp.match_id = ma.id
        '''
        df = pd.read_sql_query(query, self.system.conn)
        
        # 插入数据
        for _, row in df.iterrows():
            self.participation_tree.insert('', tk.END, values=list(row))
        
        # 更新下拉菜单
        members_df = self.system.get_all_members()
        self.part_member_combo['values'] = [f"{row['name']} (ID: {row['id']})" for _, row in members_df.iterrows()]
        
        matches_df = self.system.get_all_matches()
        self.part_match_combo['values'] = [f"{row['opponent']} ({row['date']}) ID: {row['id']}" for _, row in matches_df.iterrows()]
    
    def refresh_schedule_tab(self):
        # 清空现有数据
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)
        
        # 获取数据
        df = self.system.get_schedule()
        
        # 插入数据
        for _, row in df.iterrows():
            self.schedule_tree.insert('', tk.END, values=list(row))
        
        # 更新下拉菜单
        members_df = self.system.get_all_members()
        self.schedule_member_combo['values'] = [f"{row['name']} (ID: {row['id']})" for _, row in members_df.iterrows()]
    
    def refresh_all_tabs(self):
        self.refresh_members_tab()
        self.refresh_matches_tab()
        self.refresh_participation_tab()
        self.refresh_schedule_tab()
    
    def on_member_select(self, event):
        selection = self.members_tree.selection()
        if selection:
            item = self.members_tree.item(selection[0])
            values = item['values']
            
            # 设置按钮状态
            self.update_member_btn.config(state=tk.NORMAL)
            self.delete_member_btn.config(state=tk.NORMAL)
            
            # 填充表单
            self.member_name_var.set(values[1])
            self.member_position_var.set(values[2])
            self.member_join_date_var.set(values[3])
            self.member_experience_var.set(values[4])
            
            # 存储选中的ID
            self.selected_member_id = values[0]
        else:
            self.update_member_btn.config(state=tk.DISABLED)
            self.delete_member_btn.config(state=tk.DISABLED)
    
    def add_member(self):
        name = self.member_name_var.get()
        position = self.member_position_var.get()
        join_date = self.member_join_date_var.get()
        experience_level = self.member_experience_var.get()
        
        if not all([name, position, join_date, experience_level]):
            messagebox.showwarning("警告", "请填写所有字段")
            return
        
        try:
            self.system.add_member(name, position, join_date, experience_level)
            self.refresh_members_tab()
            self.reset_member_form()
            messagebox.showinfo("成功", f"队员 {name} 已添加")
        except Exception as e:
            messagebox.showerror("错误", f"添加队员失败: {str(e)}")
    
    def update_member(self):
        if not hasattr(self, 'selected_member_id'):
            messagebox.showwarning("警告", "请选择要更新的队员")
            return
        
        name = self.member_name_var.get()
        position = self.member_position_var.get()
        join_date = self.member_join_date_var.get()
        experience_level = self.member_experience_var.get()
        
        if not all([name, position, join_date, experience_level]):
            messagebox.showwarning("警告", "请填写所有字段")
            return
        
        try:
            self.system.update_member(self.selected_member_id, name, position, join_date, experience_level)
            self.refresh_members_tab()
            self.reset_member_form()
            messagebox.showinfo("成功", f"队员 {name} 已更新")
        except Exception as e:
            messagebox.showerror("错误", f"更新队员失败: {str(e)}")
    
    def delete_member(self):
        if not hasattr(self, 'selected_member_id'):
            messagebox.showwarning("警告", "请选择要删除的队员")
            return
        
        try:
            # 获取队员姓名用于确认
            query = "SELECT name FROM members WHERE id = ?"
            name = self.system.cursor.execute(query, (self.selected_member_id,)).fetchone()[0]
            
            confirm = messagebox.askyesno("确认", f"确定要删除队员 {name} 吗？此操作不可恢复！")
            if confirm:
                self.system.delete_member(self.selected_member_id)
                self.refresh_members_tab()
                self.reset_member_form()
                messagebox.showinfo("成功", f"队员 {name} 已删除")
        except Exception as e:
            messagebox.showerror("错误", f"删除队员失败: {str(e)}")
    
    def reset_member_form(self):
        self.member_name_var.set("")
        self.member_position_var.set("")
        self.member_join_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.member_experience_var.set("")
        self.update_member_btn.config(state=tk.DISABLED)
        self.delete_member_btn.config(state=tk.DISABLED)
        if hasattr(self, 'selected_member_id'):
            delattr(self, 'selected_member_id')
    
    def add_match(self):
        date = self.match_date_var.get()
        opponent = self.match_opponent_var.get()
        tournament = self.match_tournament_var.get()
        result = self.match_result_var.get()
        score = self.match_score_var.get()
        
        if not all([date, opponent, tournament, result]):
            messagebox.showwarning("警告", "请填写所有必填字段")
            return
        
        try:
            self.system.add_match(date, opponent, tournament, result, score)
            self.refresh_matches_tab()
            self.reset_match_form()
            messagebox.showinfo("成功", f"比赛 {opponent} 已添加")
        except Exception as e:
            messagebox.showerror("错误", f"添加比赛失败: {str(e)}")
    
    def reset_match_form(self):
        self.match_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.match_opponent_var.set("")
        self.match_tournament_var.set("")
        self.match_result_var.set("")
        self.match_score_var.set("")
    
    def record_participation(self):
        member_str = self.part_member_var.get()
        match_str = self.part_match_var.get()
        role = self.part_role_var.get()
        score = self.part_score_var.get()
        
        if not all([member_str, match_str, role]):
            messagebox.showwarning("警告", "请填写所有字段")
            return
        
        try:
            # 解析ID
            member_id = int(member_str.split('ID: ')[1])
            match_id = int(match_str.split('ID: ')[1])
            
            self.system.record_participation(member_id, match_id, role, score)
            self.refresh_participation_tab()
            self.reset_participation_form()
            messagebox.showinfo("成功", f"出战记录已添加")
        except Exception as e:
            messagebox.showerror("错误", f"记录出战失败: {str(e)}")
    
    def reset_participation_form(self):
        self.part_member_var.set("")
        self.part_match_var.set("")
        self.part_role_var.set("")
        self.part_score_var.set(5.0)
    
    def create_schedule(self):
        date = self.schedule_date_var.get()
        time_slot = self.schedule_time_var.get()
        activity = self.schedule_activity_var.get()
        member_str = self.schedule_member_var.get()
        
        if not all([date, time_slot, activity, member_str]):
            messagebox.showwarning("警告", "请填写所有字段")
            return
        
        try:
            # 解析ID
            member_id = int(member_str.split('ID: ')[1])
            
            self.system.create_schedule(date, time_slot, activity, member_id)
            self.refresh_schedule_tab()
            self.reset_schedule_form()
            messagebox.showinfo("成功", f"排班已创建")
        except Exception as e:
            messagebox.showerror("错误", f"创建排班失败: {str(e)}")
    
    def reset_schedule_form(self):
        self.schedule_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.schedule_time_var.set("")
        self.schedule_activity_var.set("")
        self.schedule_member_var.set("")
    
    def generate_performance_report(self):
        try:
            df = self.system.generate_performance_report()
            
            # 显示结果
            self.analysis_text.delete(1.0, tk.END)
            self.analysis_text.insert(tk.END, "=== 队员能力评估报告 ===\n\n")
            self.analysis_text.insert(tk.END, df.to_string(index=False))
            
            messagebox.showinfo("成功", "能力评估报告已生成")
        except Exception as e:
            messagebox.showerror("错误", f"生成报告失败: {str(e)}")
    
    def generate_match_statistics(self):
        try:
            # 获取比赛统计数据
            query = '''
                SELECT 
                    m.date,
                    m.opponent,
                    m.tournament,
                    m.result,
                    m.score,
                    COUNT(mp.id) as participants_count,
                    AVG(mp.performance_score) as avg_performance
                FROM matches m
                LEFT JOIN match_participation mp ON m.id = mp.match_id
                GROUP BY m.id
                ORDER BY m.date DESC
            '''
            df = pd.read_sql_query(query, self.system.conn)
            
            if df.empty:
                messagebox.showinfo("提示", "暂无比赛数据")
                return
            
            # 创建统计图
            fig, axes = plt.subplots(2, 2, figsize=(15, 10))
            
            # 比赛结果分布
            result_counts = df['result'].value_counts()
            axes[0, 0].pie(result_counts.values, labels=result_counts.index, autopct='%1.1f%%')
            axes[0, 0].set_title('比赛结果分布')
            
            # 比赛时间趋势
            df['date'] = pd.to_datetime(df['date'])
            df_sorted = df.sort_values('date')
            axes[0, 1].plot(df_sorted['date'], range(len(df_sorted)), marker='o')
            axes[0, 1].set_title('比赛时间线')
            axes[0, 1].set_xlabel('日期')
            axes[0, 1].set_ylabel('比赛序号')
            
            # 平均表现趋势
            if 'avg_performance' in df.columns and not df['avg_performance'].isna().all():
                axes[1, 0].plot(df_sorted['date'], df_sorted['avg_performance'], marker='s', color='green')
                axes[1, 0].set_title('平均表现趋势')
                axes[1, 0].set_xlabel('日期')
                axes[1, 0].set_ylabel('平均表现评分')
            
            # 参与人数统计
            axes[1, 1].bar(range(len(df)), df['participants_count'], color='orange', alpha=0.7)
            axes[1, 1].set_title('每场比赛参与人数')
            axes[1, 1].set_xlabel('比赛序号')
            axes[1, 1].set_ylabel('参与人数')
            
            plt.tight_layout()
            plt.savefig('match_statistics.png')
            plt.show()
            
            messagebox.showinfo("成功", "比赛统计图已生成")
        except Exception as e:
            messagebox.showerror("错误", f"生成统计图失败: {str(e)}")
    
    def generate_schedule_report(self):
        try:
            df = self.system.get_schedule()
            
            if df.empty:
                messagebox.showinfo("提示", "暂无排班数据")
                return
            
            # 创建排班可视化
            plt.figure(figsize=(12, 8))
            
            # 转换日期格式以便绘图
            df['date'] = pd.to_datetime(df['date'])
            df['activity_with_member'] = df['activity'] + '\n(' + df['assigned_member'].fillna('未分配') + ')'
            
            # 创建热力图
            pivot_data = df.pivot_table(
                index='date', 
                columns='time_slot', 
                values='activity_with_member',
                aggfunc=lambda x: x.iloc[0] if len(x) > 0 else ''
            )
            
            plt.subplot(2, 1, 1)
            plt.title('辩论队排班表')
            plt.axis('off')
            table_data = []
            for _, row in df.iterrows():
                table_data.append([
                    row['date'].strftime('%Y-%m-%d'),
                    row['time_slot'],
                    row['activity'],
                    row['assigned_member'] if pd.notna(row['assigned_member']) else '未分配'
                ])
            
            table = plt.table(cellText=table_data,
                             colLabels=['日期', '时间段', '活动', '负责人'],
                             cellLoc='center',
                             loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.5)
            
            plt.subplot(2, 1, 2)
            # 绘制时间线
            unique_dates = sorted(df['date'].unique())
            for i, date in enumerate(unique_dates):
                day_activities = df[df['date'] == date]
                for j, (_, activity) in enumerate(day_activities.iterrows()):
                    plt.barh(i, 1, left=j, label=f"{activity['time_slot']}: {activity['activity']}")
            
            plt.yticks(range(len(unique_dates)), [d.strftime('%Y-%m-%d') for d in unique_dates])
            plt.xlabel('时间段')
            plt.title('排班时间线')
            
            plt.tight_layout()
            plt.savefig('schedule_report.png')
            plt.show()
            
            messagebox.showinfo("成功", "排班表已生成")
        except Exception as e:
            messagebox.showerror("错误", f"生成排班表失败: {str(e)}")
    
    def export_data(self):
        format_type = messagebox.askquestion("导出格式", "选择导出格式:\n是 - CSV\n否 - Excel")
        format_choice = 'csv' if format_type == 'yes' else 'excel'
        self.system.export_data(format_choice)
    
    def run(self):
        self.root.mainloop()
    
    def __del__(self):
        self.system.close_connection()

def main():
    root = tk.Tk()
    app = DebateTeamApp(root)
    app.run()

if __name__ == "__main__":
    main()
