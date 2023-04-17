from tkinter import *
from pandas import DataFrame
from pandas import read_excel
from pandas import isna

class LoadFile:
    global data_dayoff
    global data_work
    
    def __init__(self,dir,load=False):
        self._dir = dir
        self.list_remarks = []
        self.list_name = []
        self.list_times = []
        self.list_dayoff = []
        self.dayoff_sum = []
        self.overtime_total = []
        self.work_day = []
        self.list_content = []
        self.list_month = [0,0,0,0,0,0,0,0,0,0,0,0]
        self.load = load
        if not(dir==''):
            self._data = read_excel(self._dir)
            self.process_data()
            
    def make_list_dayoff(self):
        len_name = 0
        for i in range(len(self._data)):
            self.list_content.append("")
            name = self._data['이름'][i]
            date = self._data['근무일자'][i]
            day = self._data['근무일명칭'][i]
            work = self._data['출근'][i]
            offwork = self._data['퇴근'][i]
            
            
            
            # 이름별 len_name # 월연차 # 배열 생성 중복제거
            if (name not in self.list_name)&(type(name)==str):
                global list_dayoff
                len_name += 1
                self.list_name.append(name)
                self.dayoff_sum.append(0)
                self.overtime_total.append(0)   
                self.work_day.append(0)   
                self.list_dayoff.append([0,0,0,0,0,0,0,0,0,0,0,0])
            
            
            # 출근 형태 : #출근 #휴무 #연장 #출장 #휴가 #기타 #반차
            remarks = ''
            hour_work = int(self._data['출근'][i][:2]) if(not isna(work)) else 0
            hour_offwork = int(self._data['퇴근'][i][:2]) if(not isna(offwork)) else 0
            if(day=='평일'):
                if(isna(work)):            
                    remarks = '휴가' if(isna(offwork)) else '누락'
                else:
                    if(isna(offwork)):
                        remarks = '누락'
                    elif(hour_work>=12):            
                        remarks = '반차'
                    elif(hour_offwork<=13):            
                        remarks = '반차'
                    elif(hour_offwork>=19):
                        remarks = '연장'
                    else:
                        remarks = '출근'
            else: # 휴무
                if(isna(work)):
                    remarks = '휴무'
                else:            
                    remarks = '연장'
                    if(isna(offwork)):
                        remarks = '누락'
            self.list_remarks.append(remarks)        
            
            # 연차 횟수 , 연장 시간 항목 추가
            self.cal_times(remarks,hour_work,hour_offwork)
                    
            # 휴가 총/월합계 # 연장 총 합계 # 근무 총 합계
            if(not isna(date)):
                if(not self.load): remarks = self._data['구분'][i]
    
                month = int(date[5:7])
                num = month-1
                index = len_name-1 # 몇 번째 사람
                # 첫번째 사람의 12월 연차 : list_dayoff[index][num] : list_dayoff[0][11] 
                if(remarks=='휴가'):
                    self.dayoff_sum[index] += 1 # 총 합계
                    self.list_dayoff[index][num] += 1 # 월 합계
                elif(remarks=='반차'):
                    self.dayoff_sum[index] += 0.5 # 총 합계
                    self.list_dayoff[index][num] += 0.5 # 월 합계
                    self.work_day[index] += 1 # 반차도 근무취급
                else: # 연장합계 # 근무 합계
                    if(remarks=='연장'):
                        self.work_day[index] += 1
                        overtime = ''
                        if(self.load):
                            overtime=hour_offwork-18 if(day=='평일') else hour_offwork-hour_work
                        else:
                            overtime=self._data['일_시간'][i]
                        self.overtime_total[index] += overtime
                    else:
                        if(day=='평일'):
                            self.work_day[index] += 1 
            
    def cal_times(self,remarks,work,off): # TODO 
        times = ''
        if(remarks=='연장'):
            times = off-18
        elif(remarks=='휴가'):
            times = 1
        elif(remarks=='반차'):
            times = 0.5
        self.list_times.append(times)
                
    def process_data(self):
        if not(self._dir==''):
            data = DataFrame({'이름':self._data['이름'],'근무일자':self._data['근무일자'],'근무일명칭':self._data['근무일명칭'],'출근':self._data['출근'],'퇴근':self._data['퇴근'],
                                 '기본':self._data['기본'],'연장':self._data['연장'],'총합':self._data['총합']})
            # 리스트 추가 # 정보 가공 
            self.make_list_dayoff()
            # 구분 컬럼 추가 # 기본 데이터 시 기본데이터 제공
            data['구분'] = self.list_remarks if self.load else self._data['구분']
            data['일_시간'] = self.list_times if self.load else self._data['일_시간']
            data['내용'] = self.list_content if self.load else self._data['내용']

            # 데이터 정리
            data = data[data['근무일자'].notnull()]
            data = data.fillna('')
            data = data.reset_index(drop=True)
            
            self.data_work = data.copy()
            self.data_dayoff = DataFrame({'이름':self.list_name,'연차':self.list_dayoff,'연차사용':self.dayoff_sum,'연장합계':self.overtime_total,'근무합계':self.work_day})
        





#################################################
