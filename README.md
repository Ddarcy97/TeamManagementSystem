# TeamManagementSystem
辩论队电子信息管理系统

项目概述

辩论队电子信息管理系统是一个基于Python的桌面应用程序，用于管理和分析辩论队的队员信息、比赛记录、出战情况和排班安排。系统提供了直观的图形用户界面，使管理员能够高效地管理辩论队的日常运营。

系统功能

队员管理：添加、编辑、删除队员信息，包括姓名、位置、加入日期和经验等级
比赛管理：记录比赛信息，包括日期、对手、锦标赛、结果和分数
出战统计：记录队员在每场比赛中的角色和表现评分
排班表管理：创建和管理队员的排班表，包括日期、时间段、活动和负责人
数据分析：生成能力评估报告、比赛统计图和排班表可视化
数据导出：将数据导出为CSV或Excel格式

系统要求

Python 3.6+
必要的Python库：
  pandas
  matplotlib
  seaborn
  sqlite3 (内置)
  tkinter (内置)

安装说明

确保已安装Python 3.6+。您可以通过以下命令检查：
      python --version
   

安装必要的Python库：
      pip install pandas matplotlib seaborn
   

将代码保存为 debate_system.py 文件

运行说明

有多种方式可以运行此程序：

方法1：命令行运行
在终端中运行：
python debate_system.py

方法2：双击运行（Windows）
保存代码为 debate_system.py
双击该文件（需要确保系统已关联Python文件）

方法3：创建可执行文件（Windows）
安装PyInstaller：
      pip install pyinstaller
   
创建可执行文件：
      pyinstaller --onefile --windowed debate_system.py
   
在 dist 文件夹中找到生成的可执行文件并运行

使用指南

启动应用：运行程序后，您将看到一个包含多个标签页的界面
队员管理：
   在"队员管理"标签页中填写队员信息
   点击"添加队员"保存信息
   选择队员后可以更新或删除
比赛管理：
   在"比赛管理"标签页中填写比赛信息
   点击"添加比赛"保存信息
出战统计：
   在"出战统计"标签页中选择队员、比赛和角色
   填写表现评分
   点击"记录出战"保存信息
排班表管理：
   在"排班表"标签页中填写排班信息
   点击"创建排班"保存信息
数据分析：
   在"数据分析"标签页中，点击相应按钮生成报告
   分析结果将显示在界面中

数据库说明

系统使用SQLite数据库存储所有数据，数据库文件为 debate_team.db，位于程序运行目录中。数据库包含以下表：

members：队员信息表
matches：比赛记录表
match_participation：出战记录表
schedule：排班表

注意：如果系统安装了R语言，系统会自动调用R进行高级统计分析。如需R语言支持，请确保已安装R并将其添加到系统PATH中。
