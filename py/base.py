from tkinter import filedialog
from tkinter import *
import os.path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import DataFrame,read_excel
from datetime import datetime
from py import load

today_year = datetime.today().year
today_month = datetime.today().month
year_var = 0
needtoload = False

base = read_excel(".\\xlsx\\employee.xlsx")

total_dayoff = []
dir = ""

def create_tk():
    def on_submit():
        global dir  
        global year_var
        selected_option = option_var.get()
        year_var = selected_option
        dir = ".\\xlsx\\{}\\total.xlsx".format(year_var)
        window.destroy()
        
    # create a new window
    window = Tk()
    window.geometry("300x150")
    window.title("연도 선택")

    # create yaers #todo
    years = []
    for year in range(2022,today_year+1):
        isdir = os.path.isdir(".\\xlsx\\{}".format(year))
        if(isdir):
            years.append(year)
    
    
    # create the option menu
    option_var = StringVar()
    option_var.set(str(today_year))  # default option
    option_menu = OptionMenu(window, option_var, *years)
    option_menu.config(width=10, borderwidth=2)
    option_menu["indicatoron"]=False
    option_menu.pack(pady=10)

    # create the submit button
    submit_button = Button(window, text="확인", width=10,command=on_submit)
    submit_button.pack(pady=20)

    window.geometry('200x120+1000+150')
    window.resizable(False,False)
    window.mainloop()

create_tk()

# 파일이 없을 경우 불러와서 생성해줌
def isfile(dir):
    global needtoload
    isfile = os.path.isfile(dir)
    if(not isfile and not year_var==0):
        dir_ = filedialog.askopenfilename()
        data = read_excel(dir_)
        write_wb = Workbook()
        write_ws = write_wb.active
        for r in dataframe_to_rows(data, index=False, header=True):
            write_ws.append(r)
        write_wb.save(dir)
        needtoload = True

def get_dayoff(): #직원정보
    
    list_name = list(base['이름'])
    list_position = list(base['직급'])
    
    data = DataFrame({'직급':list_position,'이름':list_name,'총연차':total_dayoff,'사용':"",'잔여':"",
        '1월':"",'2월':"",'3월':"",'4월':"",'5월':"",'6월':"",
        '7월':"",'8월':"",'9월':"",'10월':"",'11월':"",'12월':""})
 # type: ignore    
    
    
    data_stacked = lf.data_dayoff
    col = list(data.columns)
    for i in range(len(data)):
        name = data['이름'][i]
        for input in range(len(data_stacked)):
            if(name==data_stacked['이름'][input]):
                sum = data_stacked['연차사용'][input]
                total = float(data['총연차'][i])
                remain = total - sum
                data.loc[i,'사용'] = sum # 사용 레이블 합산 결과로 변경
                data.loc[i,'잔여'] = remain # 잔여 레이블 합산 결과로 변경
    
    return data
    
def calc_total_dayoff(): # 총 연차 계산
    join_year = list(base['입사년도'])
    join_month = list(base['입사월'])
    for i in range(len(join_year)): # 근속연수 계산
        year_work = today_year-join_year[i]
        if(year_work<1): # 1년차 일 경우
            if(join_month[i]>=today_month and join_month[i]<4): # 4월 전 이냐
                total_dayoff.append(15)
            else: # 4월 이후
                total_dayoff.append(11)
        else: # 2년차 이상
            addtion = int((year_work+1)/2)
            total_dayoff.append(15+addtion)
    

isfile(dir)
lf = load.LoadFile(dir,needtoload)    

calc_total_dayoff()

data = get_dayoff()
data2 = data.copy()
empty = data.copy()
empty2 = data.copy()
data_work = lf.data_work
emp_work = lf.data_dayoff
