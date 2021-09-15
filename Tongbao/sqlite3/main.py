import sqlite3

'''
#连接数据库，如没有则生成
conn =  sqlite3.connect('test.DB')
print('连接数据库')
cursor = conn.cursor()
#创建2886数据库结构
creat_day_report_sql = """
                    CREATE TABLE IF NOT EXISTS day_report
                    (city_name varchar(20),
                     chnl_name varchar,
                     login_no varchar(7),
                     login_name varchar,
                     op_no varchar(4),
                     op_name varchar,
                     phone_no varchar(15),
                     op_time var_char,
                     op_flow varchar(15),
                     chnl_type var_char,
                     op_bak varchar,
                     ifwd varchar(6),
                     op_count varchar(2)
                     )
"""
cursor.execute(creat_day_report_sql)

#创建县市数据表,并且写入县市数据
sql2 = """
        CREATE TABLE IF NOT EXISTS city_code_table
        (city_name varchar(20),
        city_code integer,
        login_gs varchar)
"""
cursor.execute(sql2)
sql_citycodesql = "select * from city_code_table"
city = cursor.execute(sql_citycodesql).fetchall()
if len(city)<5:
    city_list=[('盐湖营业部',1,'ja'),('芮城分公司',2,'jb'),('平陆分公司',3,'jc'),('临猗分公司',4,'jd'),('万荣分公司',5,'je'),('河津分公司',6,'jf'),
                    ('稷山分公司',7,'jg'),('垣曲分公司',8,'jh'),('绛县分公司',9,'ji'),('闻喜分公司',10,'jj'),('新绛分公司',11,'jk'),('永济分公司',12,'jl'),
                    ('夏县分公司',13,'jm'),('运城网上商城营业部',14,'jw')]
    sql_insertcitycod = "insert into city_code_table(city_name,city_code,login_gs) values(?,?,?)"
    cursor.executemany(sql_insertcitycod,city_list)
    city_list.clear()
else:
    print("数据完备")

#创建需筛选的活动的数据表结构
sql3 = """
        CREATE TABLE IF NOT EXISTS data_querysql_table
        (querysql_name varchar,sql_cmd varchar)
"""
cursor.execute(sql3)
'''
from Extkinter import *
import tkinter.ttk as ttk


class Main(Tk):
    def __init__(self):
        super().__init__()
        self.top_frame = None
        self.button_frame_1 = None
        self.button_frame_2 = None
        self.main_frame_1 = None
        self.main_frame_2 = None
        self.top_button_arr = []
        self.left_button_arr = []
        self.main_button_arr = []
        self.initialize()
        self.frame_initialize()
        self.interface_initialize()

    def initialize(self):
        self.title("数据管理系统")
        self.geometry("1000x620+%d+%d" % (self.winfo_screenwidth() / 2 - 500,
                                          self.winfo_screenheight() / 2 - 390))
        self.resizable(False, False)

    def frame_initialize(self):
        self.top_frame = Canvas(self, bg="black", height=60, width=1000, highlightthickness=0)
        self.top_frame.place(x=0, y=0)
        self.button_frame_1 = Canvas(self, bg="White", height=255, width=180, highlightthickness=0)
        self.button_frame_1.place(x=30, y=90)
        self.button_frame_2 = Canvas(self, bg="White", height=215, width=180, highlightthickness=0)
        self.button_frame_2.place(x=30, y=375)
        self.main_frame_1 = Canvas(self, bg="White", height=55, width=730, highlightthickness=0)
        self.main_frame_1.place(x=240, y=90)
        self.main_frame_2 = Canvas(self, bg="White", height=440, width=730, highlightthickness=0)
        self.main_frame_2.place(x=240, y=150)

    def interface_initialize(self):
        # 布置顶层界面
        Label(self.top_frame, text="数据管理系统", bg="black", fg="white", font=("华文细黑", 18)).place(x=30, y=13)
        top_button_1 = ExButton(self.top_frame, height=60, width=130, text="首页", font=("华文细黑", 15),
                                style="vertical_color", command=self.pass_command)
        top_button_1.set(button_list=self.top_button_arr, font_color=("White", "White"), color=("Black", "Black"),
                         active_color=("DeepSkyBlue", "Black"))
        top_button_1.place(x=300, y=0)
        top_button_2 = ExButton(self.top_frame, height=60, width=130, text="产品购买", font=("华文细黑", 15),
                                style="vertical_color", command=self.pass_command)
        top_button_2.set(button_list=self.top_button_arr, font_color=("White", "White"), color=("Black", "Black"),
                         active_color=("DeepSkyBlue", "Black"))
        top_button_2.place(x=430, y=0)
        top_button_3 = ExButton(self.top_frame, height=60, width=130, text="关于我们", font=("华文细黑", 15),
                                style="vertical_color", command=self.pass_command)
        top_button_3.set(button_list=self.top_button_arr, font_color=("White", "White"), color=("Black", "Black"),
                         active_color=("DeepSkyBlue", "Black"))
        top_button_3.place(x=560, y=0)
        top_button_4 = ExButton(self.top_frame, height=60, width=120, text="登录", font=("华文细黑", 12),
                                command=self.pass_command)
        top_button_4.set(font_color=("White", "White"), color=("Black", "Black"))
        top_button_4.place(x=880, y=0)

        # 布置按钮界面
        Label(self.button_frame_1, text="数据库", bg="white", font=("微软雅黑", 16)).place(x=30, y=10)
        left_button = ExButton(self.button_frame_1, height=40, width=180, text="数据中心", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=55)
        left_button = ExButton(self.button_frame_1, height=40, width=180, text="创建数据", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=95)
        left_button = ExButton(self.button_frame_1, height=40, width=180, text="导入数据", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=135)
        left_button = ExButton(self.button_frame_1, height=40, width=180, text="导出数据", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=175)
        left_button = ExButton(self.button_frame_1, height=40, width=180, text="设置", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=215)

        # 布置第二层按钮界面
        Label(self.button_frame_2, text="个人中心", bg="white", font=("微软雅黑", 16)).place(x=30, y=10)
        left_button = ExButton(self.button_frame_2, height=40, width=180, text="账号管理", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=55)
        left_button = ExButton(self.button_frame_2, height=40, width=180, text="我的收藏", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=95)
        left_button = ExButton(self.button_frame_2, height=40, width=180, text="我的数据", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=135)
        left_button = ExButton(self.button_frame_2, height=40, width=180, text="购买记录", font=("微软雅黑", 11),
                               command=self.pass_command)
        left_button.set(button_list=self.left_button_arr, color=("White", "White"),
                        active_color=("DeepSkyBlue", "#F0F0F0"))
        left_button.place(x=0, y=175)

        # 布置页眉
        Label(self.main_frame_1, text="数据中心", bg="White", font=("幼圆", 15)).place(x=30, y=13)
        seek_button = ExButton(self.main_frame_1, text="搜索", height=20, width=50, command=self.pass_command,
                               font=("幼圆", 12))
        seek_button.place(x=650, y=18)
        seek_entry = ttk.Entry(self.main_frame_1)
        seek_entry.place(x=480, y=18)

        # 布置主界面内容
        main_button = ExButton(self.main_frame_2, text="推荐", height=35, width=70, command=self.pass_command,
                               font=("华文细黑", 11), style="vertical_color")
        main_button.set(button_list=self.main_button_arr, active_color=("#F0F0F0", "#F0F0F0"))
        main_button.place(x=0, y=0)
        main_button = ExButton(self.main_frame_2, text="热点", height=35, width=70, command=self.pass_command,
                               font=("华文细黑", 11), style="vertical_color")
        main_button.set(button_list=self.main_button_arr, active_color=("#F0F0F0", "#F0F0F0"))
        main_button.place(x=70, y=0)
        main_button = ExButton(self.main_frame_2, text="社区", height=35, width=70, command=self.pass_command,
                               font=("华文细黑", 11), style="vertical_color")
        main_button.set(button_list=self.main_button_arr, active_color=("#F0F0F0", "#F0F0F0"))
        main_button.place(x=140, y=0)
        Label(self.main_frame_2, text="Python数据爬取\t\t2021-8-8", bg="White", font=("华文细黑", 13)).place(x=25, y=50)
        Label(self.main_frame_2, text="班级成绩数据分析\t\t2021-8-9", bg="White", font=("华文细黑", 13)).place(x=25, y=85)
        Label(self.main_frame_2, text="股票走势数据\t\t2021-8-10", bg="White", font=("华文细黑", 13)).place(x=25, y=120)
        Label(self.main_frame_2, text="硬件价格走势\t\t2021-8-10", bg="White", font=("华文细黑", 13)).place(x=25, y=155)

    def pass_command(self):
        pass


if __name__ == "__main__":
    run = Main()
    run.mainloop()


