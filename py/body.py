from tkinter import filedialog
from tkinter import *
import tkinter.ttk
from run import emp_dayoff
from run import empty_dayoff
from run import emp_work
from py import load
from py import body_tab2 as bt

# emp_dayoff 
lbl_dayoff = empty_dayoff.copy() # Label df 생성
data_dayoff = emp_work.copy()

# 불러오기 # 레이블 수치 변경
def load_file():
    dir = filedialog.askopenfilename()
    if(dir==''): return
    lf = load.LoadFile(dir,True)
    data_load_dayoff = lf.data_dayoff
    input_label_month(data_load_dayoff,True)
    bt.loadfile_event(data_load_dayoff,lf.data_work)
    
def input_label_month(data,load_bool=False): # base : 현재 데이터의 인덱스 # input : 추가할 데이터의 인덱스 # load_bool : 파일 추가 불러오기 판단
    data_dayoff = data.copy()
    col = list(emp_dayoff.columns)
    for base in range(len(emp_dayoff)):
        name = emp_dayoff['이름'][base]
        for input in range(len(data_dayoff)):
            if(name==data_dayoff['이름'][input]):
                sum = 0
                month_dayoff = data_dayoff['연차'][input]
                for i in range(len(month_dayoff)): # 월연차 합산
                    item = month_dayoff[i]
                    sum += float(item)
                    # 각 월 연차 레이블에 입력
                    txt_month_dayoff = '' if item==0 else item
                    if not(load_bool and item==0) : lbl_dayoff[col[i+5]][base].config(text=txt_month_dayoff)
                sum += float(emp_dayoff['사용'][base]) if(load_bool) else 0 # 기본데이터에 데이터 추가 if 파일을 추가 else 기본데이터 0 :(수정)
                total = float(emp_dayoff['총연차'][base])
                remain = total - sum
                lbl_dayoff['사용'][base].config(text=sum) # 사용 레이블 합산 결과로 변경
                lbl_dayoff['잔여'][base].config(text=remain) # 잔여 레이블 합산 결과로 변경

tk = Tk()
tk.title("근태관리")

# header 
header = Label(tk)
header.pack(anchor='e')

    
btn_load_file = Button(header,text="불러오기",command=load_file)
btn_load_file.pack(padx=20)

# # frame
tab=tkinter.ttk.Notebook(tk, width=1500, height=500)
tab.pack()

bt.init(tk,tab)
tab1=Frame(tk)
tab.add(tab1, text="연간연차")

# # in tab1
tab1.configure(background="white")
# col 생성
index_col = 0
for i in emp_dayoff:
    lb = Label(tab1,text=i,highlightthickness=1,highlightbackground="white", width="10",height="2",fg="white")
    if(index_col<5):
        lb.configure(background="#408080")
    else:
        lb.configure(background="#5f5f5f")
    lb.grid(row=0,column=index_col)
    index_col+=1
        
# 직원 정보 생성
for i in range(len(emp_dayoff)):
    row = emp_dayoff.loc[i]
    lbl_list = []
    j = 0; 
    for item in range(len(row)):
        lb = Label(tab1,text=row[item],highlightthickness=1,highlightbackground="white", width="10",height="2",fg="black")
        lbl_list.append(lb)
        if(j<2):
            lb.configure(background="white",highlightbackground="#daeef3")
        elif(j<5):
            lb.configure(background="#eff8fa",highlightbackground="#daeef3")
        else:
            lb.configure(background="white",highlightbackground="#eeeeee")
        lb.grid(row=i+1,column=j)
        j += 1
    lbl_dayoff.loc[i] = lbl_list # type: ignore   
input_label_month(data_dayoff)    
# in tab2 


tk.geometry('1500x500+300+100')
tk.resizable(False,False)
tk.mainloop()