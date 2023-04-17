from tkinter import *
from tkinter import messagebox
import tkinter.ttk as ttk
from pandas import DataFrame,concat 
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from run import body
import run
from py import cal

emp_base = run.emp_dayoff2.copy()
lbl_data = run.empty_dayoff2.copy()
data_work = run.data_work.copy()
emp_work = run.emp_work.copy()
check_vars = []
underlines = []
labels_work = []
name = '정건'
boolean_btn_whole = False
work_type = ""
timeHour = ""

# 생략 없이 출력 하는 방법
# pd.set_option('display.max_rows', None)
# print(data_work)

def init(tk,tab):
    tab2 = Frame(tk,background="#ffffff")
    tab.add(tab2, text="근태내역")

    # 왼쪽 프레임
    frame_left = Frame(tab2)
    frame_left.configure(background="#ffffff", height=700, width=400)

    # 왼쪽 프레임 # 근무 인원
    labelframe1 = LabelFrame(frame_left,
        background="#ffffff",
        borderwidth=0,
        font="{함초롬바탕 확장} 14 {bold}",
        foreground="#5b8ef9",
        height=600,
        labelanchor="nw",
        text=' 근무 인원')
    # 왼쪽 프레임 # 근무 인원 # col
    label2 = Label(labelframe1,
        background="#e0e9fe",
        text='직급',
        width=9)
    label2.grid(column=0, pady=1, row=0)
    label3 = Label(labelframe1,
        text='이름',
        width=9)
    label3.grid(column=1, pady=1, row=0)
    # 왼쪽 프레임 # 근무 인원 # 근무자 생성
    for i in range(len(emp_base)):
        position = emp_base.loc[i]['직급']
        name_ = emp_base.loc[i]['이름']
        label1 = Label(labelframe1,text=position,background="#e0e9fe",width=9,cursor='hand2')
        label2 = Label(labelframe1,text=name_,width=9,cursor='hand2')
        
        # 근무자 클릭 시 함수
        def select_worker(event,text):
            global name
            name=text
            total_sum() # 근태 합계
            make_work(frame10,"") # 근태 목록
            cal.make_content_text(name,"input_cal")
            
            
        label1.grid(column=0,row=i+1,pady=1)
        label2.grid(column=1,row=i+1,pady=1)
        label1.bind("<Button-1>",lambda event,text=label2['text']:select_worker(event,text))
        label2.bind("<Button-1>",lambda event,text=label2['text']:select_worker(event,text))
        lbl_data.loc[i]['직급'] = label1
        lbl_data.loc[i]['이름'] = label2
        
    # 왼쪽 프레임 # 밑줄
    underline = Frame(labelframe1)
    underline.configure(background="#5b8ef9", height=2, width=140)
    underline.place(rely=0.94, x=0, y=0)

    labelframe1.pack(ipady=5, padx=5, pady=5, side="top")

    # 왼쪽 프레임 # 근태 합계
    global label_days
    global label_dayoff
    global label_overtime
    labelframe2 = LabelFrame(frame_left)
    labelframe2.configure(
        background="#ffffff",
        borderwidth=0,
        font="{함초롬바탕 확장} 14 {bold}",
        foreground="#8670a0",
        labelanchor="nw",
        text=' 근태 합계')
    label15 = Label(labelframe2)
    label15.configure(
        background="#e4e0eb",
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='근태(일수)',
        width=9)
    label15.grid(column=0, row=0)
    label_days = Label(labelframe2)
    label_days.configure(
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='--',
        width=9)
    label_days.grid(column=1, row=0)
    label17 = Label(labelframe2)
    label17.configure(
        background="#e4e0eb",
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='연차(공제)',
        width=9)
    label17.grid(column=0, row=1)
    label_dayoff = Label(labelframe2)
    label_dayoff.configure(
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='--',
        width=9)
    label_dayoff.grid(column=1, row=1)
    label91 = Label(labelframe2)
    label91.configure(
        background="#e4e0eb",
        font="TkDefaultFont",
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='연장(시간)',
        width=9)
    label91.grid(column=0, row=2)
    label_overtime = Label(labelframe2)
    label_overtime.configure(
        height=2,
        highlightbackground="#ded7e6",
        highlightthickness=1,
        text='--',
        width=9)
    label_overtime.grid(column=1, row=2)

    # 왼쪽 프레임 # 밑줄2
    underline2 = Frame(labelframe2)
    underline2.configure(background="#8670a0", height=2, width=145)
    underline2.place(anchor="nw", rely=0.97, x=0, y=0)

    labelframe2.pack(padx=5, pady=10, side="bottom")
    frame_left.pack(anchor="n", ipadx=20,ipady=150, side="left")

    # 중앙 캘린더 # 달력
    cal.init(tab2)
    cal.make_content_text(name,"input_cal")
    
    separator = ttk.Separator(tab2, orient="vertical")
    separator.pack(fill="y", ipady=200, padx=10,pady=10, side="left")

    # 오른쪽 프레임
    frame_right = Frame(tab2)
    frame_right.configure(
        background="#ffffff", height=550, width=700)
    
    # 선택된 근무자 레이블
    global label_worker
    label_worker = Label(frame_right, text=name, background="#ffffff", font="14")
    label_worker.place(anchor="nw",x=50,y=10)
    

    # 전체보기 버튼
    btn_whole = Button(frame_right,text='전체보기',command=lambda : make_work(frame10))
    btn_whole.place(anchor="ne", x=430, y=10)
    
    # 수정 버튼
    btn_edit = Button(frame_right,text='수정',command=edit_btn_action)
    btn_edit.place(anchor="ne", x=480, y=10)
    
    # 저장 버튼
    btn_save = Button(frame_right,text='저장',command=btn_save_action)
    btn_save.pack(anchor="ne", padx=10, pady=10, side="top")

    # 프레임 # 스크롤 + 표
    frame9 = Frame(frame_right)
    frame9.configure(background="#ffffff")

    # 스크롤
    scrollbar = Scrollbar(frame9)
    scrollbar.configure(orient="vertical")
    scrollbar.pack(fill="y", side="right" , padx = 20)
    # 스크롤 가능 위젯
    global canvas
    canvas = Canvas(frame9, width=450,yscrollcommand=scrollbar.set)
    scrollbar.config(command=canvas.yview)
    def scroll_mouse(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", scroll_mouse)
    # 표
    global frame10
    frame10 = Frame(canvas,bg="#ffffff")
    frame10.configure(width=900)
    frame10.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
    
    lbl_col1 = Label(frame10,background="#6889ae",foreground="#fcfdfe",text='선택',width=7)
    lbl_col1.grid(column=0, padx=1, row=0)
    
    lbl_col0 = Label(frame10,background="#6889ae",foreground="#ffffff",text='이름',width=9)
    lbl_col0.grid(column=1, padx=1, row=0)
    
    lbl_col = Label(frame10,background="#6889ae",foreground="#ffffff",text='일자',width=11)
    lbl_col.grid(column=2, padx=1, row=0)
    
    lbl_col2 = Label(frame10,background="#6889ae",foreground="#ffffff",text='근태항목',width=9)
    lbl_col2.grid(column=3, padx=1, row=0)
    
    lbl_col3 = Label(frame10,background="#6889ae",foreground="#ffffff",text='일_시간')
    lbl_col3.grid(column=4, padx=1, row=0)
    
    lbl_col4 = Label(frame10,background="#6889ae",foreground="#ffffff",text='내용',width=14)
    lbl_col4.grid(column=5, padx=1, row=0)
    
    canvas.create_window((0, 0), window=frame10,width=450, anchor='nw')
    canvas.pack(side="right")
    frame9.pack(side="top")
    frame_right.pack(side="left",anchor='nw',ipadx=40)
    total_sum()
    make_work(frame10,'')
    cal.make_content_text(name)

# 클릭 시 근태합계 데이터 뿌리기
def total_sum():
    for i in range(len(emp_work)):
        if(name==emp_work['이름'][i]):
            text_days = emp_work['근무합계'][i] if(not emp_work['근무합계'][i]==0 ) else '--'
            text_dayoff = emp_work['연차사용'][i] if(not emp_work['연차사용'][i]==0 ) else '--'
            text_overtime = emp_work['연장합계'][i] if(not emp_work['연장합계'][i]==0 ) else '--'
            
            label_days.config(text=text_days)
            label_dayoff.config(text=text_dayoff)
            label_overtime.config(text=text_overtime)

# 이름에 맞는 레이블 생성 함수
def make_work(frame10,con='whole'):
    global boolean_btn_whole
    month = cal.month
    
    #scrollbar 올리기
    canvas.yview_moveto(0)
    # 선택된 근무자 레이블 이름 변경
    label_worker.config(text=name)
    
    if(con=='whole'):
        boolean_btn_whole = False if(boolean_btn_whole) else True
    else:
        boolean_btn_whole = False
    i = 0
    for j,array_label in enumerate(labels_work): # 테이블 레이블 # 밑줄 모두 삭제
        underlines[j].destroy()
        for label in array_label:
            label.destroy()
    labels_work.clear()
    check_vars.clear()
    underlines.clear()
    for num , row in data_work.iterrows():
        condition_name = name==row['이름']
        condition_month = (month==int(row['근무일자'][5:7]))
        condition_work = (not row['구분']=='출근') and (not row['구분']=='휴무')
        condition = condition_name and condition_work and condition_month if(con=='') else condition_name and condition_month
        condition = condition_name and condition_work and condition_month if(not boolean_btn_whole) else condition_name and condition_month
        if(condition):
            i += 1
            checkbox_var = BooleanVar()
            checkbox = Checkbutton(frame10, text="", variable=checkbox_var,background='#ffffff')
            
            label_name = Label(frame10,text=name,background='#ffffff')
            label_date = Label(frame10,text=row['근무일자'],background='#ffffff')
            label_time = Label(frame10,text=row['구분'],background='#ffffff')
            label_timesHour = Label(frame10,text=row['일_시간'],background='#ffffff')
            label_content = Label(frame10,text=row['내용'],background='#ffffff')
            
            checkbox.grid(column=0,row=i)
            label_name.grid(column=1,row=i)
            label_date.grid(column=2,row=i)
            label_time.grid(column=3,row=i)
            label_timesHour.grid(column=4,row=i)
            label_content.grid(column=5,row=i)
            
            underline3 = Frame(frame10,background="#6889ae", height=1, width=420)
            underline3.place(anchor="nw",x=10, y=19+25*i)
            
            underlines.append(underline3)
            check_vars.append(checkbox_var)
            labels_work.append([checkbox,label_name,label_date,label_time,label_timesHour,label_content])      
            
# 데이터 추가 로드 시 데이터 합산 및 입력
def loadfile_event(emp_work_,data_work_): #emp_work_ : 로드한 합계 데이터 data_work_: 로드한 모든 데이터
    # 근태내역 필요 데이터 합산
    sum_data_emp_work(emp_work_)
    # 데이터 합산
    sum_data_work(data_work_)
    
    # 데이터 갱신
    cal.data_work = data_work
    
    # 데이터 입력
    total_sum()
    make_work(frame10,'')
    
# 데이터 합산 함수    
def sum_data_emp_work(emp_work_):
    for i , row_ in emp_work_.iterrows():
        name_ = row_['이름']
        for j , row in emp_work.iterrows():
            name = row['이름']
            if(name==name_):
                # 월별 연차 합산
                for k in range(len(row['연차'])):
                    row['연차'][k] += row_['연차'][k]
                # 연차사용 합산
                emp_work['연차사용'][j]=row['연차사용'] + row_['연차사용']
                # 연장합계 합산
                emp_work['연장합계'][j]=row['연장합계'] + row_['연장합계']
                # 근무합계 합산
                emp_work['근무합계'][j]=row['근무합계'] + row_['근무합계']

# 데이터 합산 함수    
def sum_data_work(data_work_):
    global data_work
    data_work_.reset_index(inplace=True,drop=True)
    for i in range(len(data_work_)):
        row_ = data_work_.iloc[i]
        name_ = row_['이름']
        row_df = DataFrame([list(data_work_.iloc[i])],columns=list(data_work_.columns))
        index = (data_work[data_work['이름']==name_].index[-1]) +1 #type:ignore
        data_work = concat([data_work.iloc[:index], row_df, data_work.iloc[index:]], ignore_index = True)  
    
# 수정 버튼 액션
def edit_btn_action():
    i = 0
    for check_var in check_vars:
        if(check_var.get()):
            i += 1
    if(not i==0):
        global work_type
        global timeHour
        work_type = ""
        timeHour = ""
        create_frame_edit()
    else:
        messagebox.showinfo("Alert", "수정할 목록을 체크해주세요")
    
    
# 수정 프레임 생성
def create_frame_edit():
    edit_tk = Tk()
    edit_tk.title("수정")
    frame = Frame(edit_tk,width=200,height=100)
    frame.pack(side='top',pady=10)
    frame.grid_propagate(False)
    frame.grid_anchor('center')
    
    lbl_work = Label(frame,text='근태항목',width=12)
    lbl_work.grid(column=0,row=0,padx=1,pady=5)
    
    lbl_timeHour = Label(frame,text='일_시간',width=12)
    lbl_timeHour.grid(column=0,row=1, pady=5)
    
    lbl_content = Label(frame,text='기타내용',width=12)
    lbl_content.grid(column=0,row=2, pady=5)
    
    # 콤보 박스 # 근태항목
    list_work = ["출근","휴가","반차","출장","연장","누락","기타"]
    variable_work = StringVar(frame)
    variable_work.set("근태항목")

    opt_work = OptionMenu(frame, variable_work, *list_work)
    opt_work.config(width=10, bg="#ffffff",borderwidth=1)
    opt_work["indicatoron"]=False
    opt_work.grid(column=1,row=0, pady=5)
    
    def callback_work(*args): # 클릭 이벤트
        global work_type
        work_type = (variable_work.get())

    variable_work.trace("w", callback_work)   
    
    # 콤보 박스 # 근태항목
    list_timeHour = [0.5,1,2,3,4,5,6]
    variable_timeHour = StringVar(frame)
    variable_timeHour.set("일수(시간)")

    opt_timeHour = OptionMenu(frame, variable_timeHour, *list_timeHour)
    opt_timeHour.config(width=10, bg="#ffffff",borderwidth=1)
    opt_timeHour["indicatoron"]=False
    opt_timeHour.grid(column=1,row=1, pady=5)
    
    def callback_timeHour(*args): # 콤보박스 클릭 이벤트
        global timeHour
        timeHour = (variable_timeHour.get())

    variable_timeHour.trace("w", callback_timeHour)   
    
    txt_content = Entry(frame,width=11)
    txt_content.grid(column=1,row=2, pady=5)
    
    # 확인 버튼 액션
    btn_confirm = Button(edit_tk,text="확인",command=lambda:edit_confirm_action(edit_tk,work_type,timeHour,txt_content.get()))
    btn_confirm.pack(side="bottom",anchor="s",pady=10)
    
    
    edit_tk.geometry('200x170+1000+150')
    edit_tk.resizable(False,False)
    edit_tk.mainloop()
    
# 수정 프레임 확인 버튼 액션
def edit_confirm_action(edit_tk,work_type_,timeHour_,content_):
    edit_tk.destroy()
    if(work_type_=='' and timeHour_=='' and content_==''): # 공란일시 메세지 알림
        messagebox.showinfo("Alert","수정할 항목을 선택/기입하세요")
        create_frame_edit()
    else:
        for i,check_var in enumerate(check_vars): # 체크한 데이터만 뽑아오기
            if(check_var.get()==True):
                name_ = labels_work[i][1]["text"]
                date = labels_work[i][2]["text"]
                work_type = labels_work[i][3]["text"]
                
                # 데이터 수정 # 근태목록 데이터
                condition = (data_work['이름']==name_) & (data_work['근무일자']==date)
                data_work.loc[condition,'일_시간']=timeHour_
                data_work.loc[condition,'내용']=content_
                if(work_type_==''):
                    data_work.loc[condition,'구분']=work_type
                else:
                    data_work.loc[condition,'구분']=work_type_
                if work_type_ == '휴가' : data_work.loc[condition,'일_시간']=1
                if work_type_ == '반차' : data_work.loc[condition,'일_시간']=0.5
        # 데이터 수정 # 근태합계 데이터
        refresh_total_sum()
        cal.data_work = data_work            
        # 데이터 기입
        total_sum() # 근태 합계
        make_work(frame10,"") # 근태 목록
        cal.make_content_text(name,"input_cal") # 달력 근태현황
        body.input_label_month(emp_work) # 탭1 연간연차
        

# 근태 합계 수정 함수
def refresh_total_sum():
    for i in range(len(emp_work)):
        dayoff= []
        dayoff_year = 0
        name_ = emp_work.iloc[i]['이름']
        # 월 연차합산
        for j in range(1,13):
            month = str(j).zfill(2)
            dayoff_month = 0
            for k in range(len(data_work)): 
                condition_month = ( 
                    (data_work['이름'][k]==str(name_)) and
                    ('2023/'+month) in (data_work["근무일자"][k]) and
                    ((data_work['구분'][k]=='휴가') or (data_work['구분'][k]=='반차')) 
                )
                if(condition_month):
                    dayoff_month += float(data_work['일_시간'][k])
            dayoff_year += float(dayoff_month)
            dayoff.append(dayoff_month)
        # 총 근무합산
        condition_workdays = (
            (data_work.이름==str(name_))&
            (data_work.근무일명칭=='평일')&
            (data_work.구분!='휴가')
        )
        # 월 연장합산
        condition_overtime = ( 
            (data_work.이름==str(name_))&
            (data_work.구분=='연장')
        )
        emp_work.at[i,'연차'] = dayoff
        emp_work.at[i,'연차사용'] = dayoff_year
        emp_work.at[i,'근무합계'] = len(data_work.loc[condition_workdays,:])
        emp_work.at[i,'연장합계'] = data_work.loc[condition_overtime,'일_시간'].copy().sum()
            
def btn_save_action():
    path = ".\\xlsx\\{}\\total.xlsx".format(run.year_var)
    write_wb = Workbook()
    write_ws = write_wb.active
    for r in dataframe_to_rows(data_work, index=False, header=True):
        write_ws.append(r)
    write_wb.save(path)
    
