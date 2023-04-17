from tkinter import *
from datetime import datetime
import calendar
import holidays
from run import data_work
from py import body_tab2 as bt

# 데이터


# 달력 만들기
lastindex = len(data_work)-2
year = int(data_work['근무일자'][lastindex][0:4]) if lastindex>0 else datetime.today().year
month = int(data_work['근무일자'][lastindex][5:7]) if lastindex>0 else datetime.today().month

cal_col = ['SUN','MON','TUE','WED','THU','FRI','SAT']

name = ''
list_lbl_days = []
list_lbl_content = []
list_content_text= []

# 월 리스트
list_month = ["1","2","3","4","5","6","7","8","9","10","11","12"]


def init(tk):
    
    # 달력 안에 월에 맞는 날짜 , 내용 생성
    frame = Frame(tk,background="#ffffff")
    frame.pack(side='left',anchor='nw', padx=10, ipadx=20)

    # 년도 레이블 생성
    label_year = Label(frame,text=str(year),bg="#ffffff",font="20")
    label_year.place(x=50,y=17)
    
    # 월 선택 메뉴
    variable = StringVar(frame)
    variable.set(str(month))

    opt = OptionMenu(frame, variable, *list_month)
    opt.config(width=10, bg="#ffffff",borderwidth=0)
    opt["indicatoron"]=False
    opt.pack(side="top",anchor="nw",padx=230,pady=10)

    def callback(*args): # 월 변경 클릭 이벤트
        global month
        month = int(variable.get())
        bt.make_work(bt.frame10,'')
        make_content_text(name,"input_cal")

    variable.trace("w", callback)   

    # 달력 테이블 생성
    global frame_table
    frame_table = Frame(frame,background="#ffffff")
    frame_table.pack(side='left',anchor='nw')

    for c in range(7):
        label = Label(frame_table,text=cal_col[c],width=8,padx=10,pady=10,bg="#ffffff")
        label.grid(column=c,row=0)
        
# 달력 날짜 , 내용 기입
def input_label_day(month,frame_table):
    global list_lbl_days
    global list_lbl_content
    
    if(not list_lbl_days==[]):
        # 레이블 삭제
        for r in range(0,6):
            for label in list_lbl_days[r]:
                label.destroy()
    
    list_lbl_days.clear()
    list_lbl_content.clear()
    
    c = calendar.TextCalendar(calendar.SUNDAY)
    list_days = c.monthdayscalendar(year,month)

    # 배열 길이 6개로 고정
    while(len(list_days)<6):
        list_days.append([0,0,0,0,0,0,0])
        
    # 레이블 배열 생성
    list_lbl_days =[[],[],[],[],[],[],[]]
    list_lbl_content =[[],[],[],[],[],[],[]]
    

    for r in range(1,7):
        for c in range(7):
            color = "#efefef"
            fg_color = "black"
            txt_content = ""
            txt_date=list_days[r-1][c]
            if(txt_date==0): # 날짜가 없는 칸은 비움
                color = "#ffffff"
                txt_date=""
            for list in list_content_text: # 알맞는 날짜에 내용 기입
                if(txt_date==list[0]):
                    txt_content = list[1]
            if c==0: fg_color='red'
            elif c==6: fg_color='blue'
            holiday_name = check_holiday(txt_date)
            if not holiday_name=='':
                fg_color='red'
                txt_content=holiday_name
            
            labelframe = LabelFrame(frame_table,
                background=color,
                borderwidth=0,
                labelanchor="n",
                text=txt_date,
                padx=15,
                fg=fg_color,
                pady=10)
            labelframe.grid(column=c,row=r,padx=1,pady=1)
            list_lbl_days[r-1].append(labelframe)
            
            label_content = Label(labelframe,text=txt_content,background=color,width=8,foreground=fg_color)
            label_content.pack()
    
# 달력 내용 생성
def make_content_text(name_,con=""):
    global name
    name = name_
    list_content_text.clear()
    for num , row in data_work.iterrows():
        if((name_==row['이름']) and (not row['구분']=='출근') and (not row['구분']=='휴무') and month==int(row['근무일자'][5:7])):
            date = int(row['근무일자'][8:10])
            list_content_text.append([date,row['구분']])
    if(con=="input_cal"):
        input_label_day(month,frame_table)
        
# holidays
def check_holiday(date):
    if date=="" : return None
    datetime_ = '{}-{}-{}'.format(month,date,year)
    isholiday= holidays.KR().get(datetime_) if not holidays.KR().get(datetime_)==None else ''
    if "Alternative holiday" in isholiday: isholiday="Alternative holiday"
    holiday_names = {
        "":"",
        "New Year's Day":"신정",
        "The day preceding of Lunar New Year's Day":"설날",
        "Lunar New Year's Day":"설날",
        "The second day of Lunar New Year's Day":"설날",
        "Independence Movement Day":"삼일절",
        "Labour Day":"근로자의날",
        "Children's Day":"어린이날",
        "Birthday of the Buddha":"석가탄신일",
        "Memorial Day":"현충일",
        "Liberation Day":"광복절",
        "The day preceding of Chuseok":"추석",
        "Chuseok":"추석",
        "The second day of Chuseok":"추석",
        "National Foundation Day":"개천절",
        "Hangeul Day":"한글날",
        "Christmas Day":"크리스마스",
        "Alternative holiday":"대체공휴일"
    }
    return holiday_names[isholiday]
