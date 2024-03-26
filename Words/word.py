#@FMC 2023 Autumn at RUC

#多行注释代码块：Ctrl+\
import pandas as pd
import datetime
from tkinter import *
from tkinter import ttk
import sys
import ctypes
import re
#告诉操作系统使用程序自身的dpi适配
ctypes.windll.shcore.SetProcessDpiAwareness(1)
#获取屏幕的缩放因子
ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)

class Word_list: #目前最大支持创建3699个list
    def __init__(self, word_table, label, additional_info, list_info):
        self.word_table = word_table  #存储单词表
        self.label_num = label #单词表编号
        self.additional_info = additional_info #包含单词遗忘次数、意群（含备注）
        self.word_num = word_table.shape[0] #获取行数
        self.recur_days = [2,2,3,4,9]  #几天之后复习
        self.repeat_times = list_info.iat[0, 0] #已经复习的次数，通过Recite.save_record函数保存
        self.curr_review_index = list_info.iat[0, 2]
        self.finishing_time = datetime.datetime.strptime(str(list_info.iat[0, 1]), '%Y-%m-%d') #记录本次背诵完成的时间，通过Recite.save_record函数更新，存储到list_info
        if(self.repeat_times >= len(self.recur_days)): #复习完成
            self.next_time = datetime.datetime.today() + datetime.timedelta(days = 1) #往后推一天，则后续永远不会再复习到
        elif(self.repeat_times == -1): #尚未开始复习
            self.next_time = self.finishing_time
        else:
            self.next_time = self.finishing_time + datetime.timedelta(days = self.recur_days[self.repeat_times]) #对下一次复习计划进行设定


class Recite:
    def __init__(self, plan_num): #num表示一个list中的单词数目
        self.num = plan_num
        self.present_list_label = 0

        self.word = pd.read_excel("E:/Words/data.xlsx") #本部分用于读入单词表，构建word_list对象
        self.add_info = pd.read_excel("E:/Words/add_info.xlsx") #单词的额外信息
        self.history = pd.read_excel("E:/Words/history.xlsx") #历史记录
        self.list_info = pd.read_excel("E:/Words/list_info.xlsx") #list的信息，一行代表一整个word_list

        self.curr_list = self.history.iat[0,0]
        self.curr_word_index = self.history.iat[0,1]
        self.last_list = 0
        self.last_word_index = 0

        self.list_num = self.word.shape[0]//self.num + 1
        start_index = 0
        self.word_list = []
        for i in range(self.list_num):
            temp = Word_list(self.word[min(start_index, self.word.shape[0] - 1) : min(start_index + self.num, self.word.shape[0] - 1)], i+1, self.add_info[min(start_index, self.word.shape[0]) : min(start_index+self.num+1, self.word.shape[0])], self.list_info[i:i+1])
            self.word_list.append(temp)  #word_list中包含多个Word_list对象
            start_index += self.num

    def display_word(self):
        temp_word = str(self.word_list[self.curr_list].word_table.iat[self.curr_word_index, 0]) #第0列表示单词,text表示label的具体内容
        label['text'] = temp_word
        label['font'] = ('Times New Roman', 36)
        temp_index = sent_data.loc[sent_data['word'] == temp_word].index[0] #确定单词对应的句子的索引
        sysn['text'] = sent_data.iat[temp_index, 1]
        sent['text'] = sent_data.iat[temp_index, 2]
        label_3['text'] = 'List ' + str(self.curr_list + 1) + '-' + str(self.curr_word_index + 1)
        label_3['font'] = ('Times New Roman', 10)
        self.last_list = self.curr_list
        self.last_word_index = self.curr_word_index

        if(self.curr_word_index < self.word_list[self.curr_list].word_table.shape[0] - 1): #仍在老的word list中
            self.curr_word_index += 1
        else:
            if(self.curr_list < self.list_num - 1): #进入到一个新的word list
                if(self.word_list[self.curr_list].repeat_times == -1):
                    self.word_list[self.curr_list].repeat_times += 1  #将首次学习完成的word_list的状态置为复习状态（repeat_times = -1代表初学）
                    self.word_list[self.curr_list].finishing_time = datetime.date.today()
                    self.list_info.iloc[self.word_list[self.curr_list].label_num - 1, 1] = datetime.date.today().strftime('%Y-%m-%d') #更新原表日期
                    self.list_info.iloc[self.word_list[self.curr_list].label_num - 1, 0] = self.list_info.iat[self.word_list[self.curr_list].label_num - 1, 0] + 1
                self.curr_list += 1
                self.curr_word_index = 0  #重置单词索引
            else:
                label['text'] = '单词背完啦！'
                label['font'] = ('宋体', 20)

        #出现按钮
        button.grid_forget()
        button_1.grid(row=3, padx=(46,180), pady=15)
        button_2.grid(row=3, padx=(180,46), pady=15)
        button_3.grid(row=5)
        button_4.grid(row=6)
        button_5.grid(row=4)
        button_6.place(relx=0, rely=1, anchor=SW)
        my_window.unbind("<space>")

        #仅出现单词，认为尚未背诵，故索引更新为旧索引
        self.history.iloc[0,0] = self.last_list
        self.history.iloc[0,1] = self.last_word_index

    def save_record(self):
        self.history.to_excel("E:/Words/history.xlsx", index = False)  #别忘记保存记录！
        self.add_info.to_excel("E:/Words/add_info.xlsx", index = False)   
        self.list_info.to_excel("E:/Words/list_info.xlsx", index = False)
        sys.exit(0)

    def record_word(self): #记录遗忘次数
        self.add_info.iloc[self.last_list*self.num + self.last_word_index, 0] = self.add_info.iat[self.last_list*self.num + self.last_word_index, 0] + 1   #实时更新


fmc_recite = Recite(88)

my_window = Tk() #主窗口
my_window.title('快乐背单词神器')
my_window.tk.call('tk', 'scaling', ScaleFactor/75) #设置程序缩放

sentence_window = Toplevel()
sentence_window.title('近义词和例句展示')
sentence_window.tk.call('tk', 'scaling', ScaleFactor/75)
sentence_window.attributes("-toolwindow", 1) #隐藏右上角三大金刚键

label = Label(my_window) #显示单词
label.grid(row=0)
label_2 = Label(my_window) #显示单词释义
label_2.grid(row=1)
label_3 = Label(my_window) #显示当前学习的word_list
label_3.place(rely=1.0, relx=1.0, x=0, y=0, anchor=SE)

sysn = Message(sentence_window, width=500)
sysn['font'] = ('微软雅黑', 14)
sysn.grid(row = 0)
sent = Message(sentence_window, width=500)
sent['font'] = ('微软雅黑', 14)
sent.grid(row = 2)
sent_separator = ttk.Separator(sentence_window, orient='horizontal')
sent_separator.grid(row = 1, sticky='ew')

sent_data = pd.read_excel("E:/Words/sentence_data.xlsx") #包含近义词和例句的词库
sent_data['sentence'] = sent_data['sentence'].str.replace('_x000D_', '') #解决换行符问题
sent_data['synonyms'] = sent_data['synonyms'].str.replace('_x000D_', '') #解决换行符问题
sent_data = sent_data.fillna('Nothing~~') #处理NaN值
def remove_chinese(text):
    chinese_pattern = re.compile('[\u4e00-\u9fa5]')
    return chinese_pattern.sub('', text)

sent_data['sentence'] = sent_data['sentence'].apply(remove_chinese) #在数据框的指定列上应用这个函数
sent_data['sentence'] = sent_data['sentence'].str.replace('\n', ' ') #解决空白行问题

def clear_mean():
    label_2['text'] = ' '
    label_2['font'] = ('微软雅黑', 14)

def remem_button(): #单击"有印象后弹出"
    my_window.unbind("<Left>") #临时禁用左热键
    my_window.unbind("<Right>") #临时禁用右热键
    my_window.unbind("<Control-s>") #临时禁用保存热键

    label_2['text'] = str(fmc_recite.word_list[fmc_recite.last_list].word_table.iat[fmc_recite.last_word_index, 1])
    label_2['font'] = ('微软雅黑', 14)

    temp_1 = Button(my_window, text = "正确，继续背诵 ↑", font=("微软雅黑",10),fg="blue", command = lambda:[clear_mean(), fmc_recite.display_word(), temp_1.grid_forget(), temp_2.grid_forget(), my_window.unbind("<Up>"), my_window.unbind("<Down>"), my_window.bind("<Left>", lambda event: button_1.invoke()), my_window.bind("<Right>", lambda event: button_2.invoke()), my_window.bind("<Control-s>", lambda event: button_3.invoke())]) #恢复绑定左右热键、保存键
    temp_2 = Button(my_window, text = "错误，记录并继续背诵 ↓", font=("微软雅黑",10),fg="blue", command = lambda:[fmc_recite.record_word(), clear_mean(), fmc_recite.display_word(), temp_1.grid_forget(), temp_2.grid_forget(), my_window.unbind("<Up>"), my_window.unbind("<Down>"), my_window.bind("<Left>", lambda event: button_1.invoke()), my_window.bind("<Right>", lambda event: button_2.invoke()), my_window.bind("<Control-s>", lambda event: button_3.invoke())])
    temp_1.grid(row=7, padx=(46,260), pady=15)
    temp_2.grid(row=7, padx=(260,46), pady=15)

    button_1.grid(row=3, padx=(46,180), pady=15) #重新显示按钮
    button_2.grid(row=3, padx=(180,46), pady=15)
    button_3.grid(row=5)

    fmc_recite.history.iloc[0,0] = fmc_recite.curr_list
    fmc_recite.history.iloc[0,1] = fmc_recite.curr_word_index

    #设置热键
    my_window.bind("<Up>", lambda event: temp_1.invoke()) #正确，继续背诵
    my_window.bind("<Down>", lambda event: temp_2.invoke()) #错误，记录并继续背诵

def not_remem_button():
    my_window.unbind("<Left>") #临时禁用左热键
    my_window.unbind("<Right>") #临时禁用右热键
    my_window.unbind("<Control-s>") #临时禁用保存热键   

    label_2['text'] = str(fmc_recite.word_list[fmc_recite.last_list].word_table.iat[fmc_recite.last_word_index, 1])
    label_2['font'] = ('微软雅黑', 14)

    fmc_recite.record_word() #点击忘记，意思就显示出来，从而自动记录（本操作不能放到temp_3按钮中！）
    temp_3 = Button(my_window, text = "记住了，下一个吧<Enter>", font=("微软雅黑",10),fg="blue", command = lambda:[clear_mean(), fmc_recite.display_word(), temp_3.grid_forget(), my_window.unbind("<Return>"), my_window.bind("<Left>", lambda event: button_1.invoke()), my_window.bind("<Right>", lambda event: button_2.invoke()), my_window.bind("<Control-s>", lambda event: button_3.invoke())])
    temp_3.grid(row=7, pady=15)

    button_1.grid(row=3, padx=(46,180), pady=15) #重新显示按钮
    button_2.grid(row=3, padx=(180,46), pady=15)
    button_3.grid(row=5)

    fmc_recite.history.iloc[0,0] = fmc_recite.curr_list
    fmc_recite.history.iloc[0,1] = fmc_recite.curr_word_index

    #设置热键
    my_window.bind("<Return>", lambda event: temp_3.invoke()) #记住了，下一个吧

def open_review(): #每次要复习的是当前时间点之前的所有时间组对象包含的word_list
    review_window = Toplevel() #复习任务窗口
    review_window.title('今日复习任务')
    review_window.tk.call('tk', 'scaling', ScaleFactor/75) #设置窗口显示缩放

    #today_review_list = []  #迄今为止要复习的索引，每次取第一个作为当前要复习的列表，复习完之后则立即弹出
    today_review_frame = pd.DataFrame({'list_index':[], 'next_time':[]})
    for i in range(len(fmc_recite.word_list)):
        if(fmc_recite.word_list[i].next_time <= datetime.datetime.today() and fmc_recite.word_list[i].repeat_times > -1):
            #today_review_list.append(i)
            today_review_frame.loc[len(today_review_frame.index)] = [i, fmc_recite.word_list[i].next_time]
    today_review_frame.sort_values('next_time', inplace = True) #直接在原表上排序
    today_review_list = today_review_frame['list_index'].values.tolist()
    today_review_list = [int(i) for i in today_review_list]  #全部转化为整型

    r_label_0 = Label(review_window)
    r_label_0.grid(row=0, column=1)
    r_label_0['text'] = '还没复习的Word List:' + str([i + 1 for i in today_review_list])
    r_label_1 = Label(review_window) #显示单词
    r_label_1.grid(row=1, column=1)
    r_label_2 = Label(review_window) #显示单词释义
    r_label_2.grid(row=2, column=1)
    text_review = Text(review_window, height=2, width=30)
    text_review.configure(font=("微软雅黑", 12))

    def start():
        if(len(today_review_list) != 0):
            curr_list = today_review_list[0]
            curr_review_index = fmc_recite.word_list[curr_list].curr_review_index
            curr_total_index = curr_list*fmc_recite.num + curr_review_index

            temp_word_1 = str(fmc_recite.word_list[curr_list].word_table.iat[curr_review_index, 0])
            r_label_1["text"] = temp_word_1
            r_label_1["font"] = ('Times New Roman', 36)
            temp_index_1 = sent_data.loc[sent_data['word'] == temp_word_1].index[0] #确定单词对应的句子的索引
            sysn['text'] = sent_data.iat[temp_index_1, 1]
            sent['text'] = sent_data.iat[temp_index_1, 2]
            text_review.delete(1.0, "end") #先删除原有内容
            text_review.insert(1.0, fmc_recite.add_info.iat[curr_total_index, 1]) #再显示新内容
            r_button_0.grid_forget()
            r_button_1.grid(row=4, column=1)
            r_button_4.grid(row=5, column=1)
            text_review.grid(row=3, columnspan=3)
        else:
            r_label_1["text"] = 'Good boy! No word list to review today!'
            r_label_1["font"] = ('Times New Roman', 36)
            r_button_0.grid_forget()

        return(review_window.unbind("<space>"), review_window.bind("<Return>", lambda event: r_button_1.invoke()))
        
    def check():
        curr_list = int(today_review_list[0])
        curr_review_index = fmc_recite.word_list[curr_list].curr_review_index

        r_label_2["text"] = str(fmc_recite.word_list[curr_list].word_table.iat[curr_review_index, 1])
        r_label_2["font"] = ('微软雅黑', 14)
        r_label_2.grid(row=2, column=1)

        r_button_1.grid_forget()
        r_button_4.grid_forget()
        r_button_2.grid(row=4, column=0)
        r_button_3.grid(row=4, column=2)

        return(review_window.unbind("<Return>"), review_window.bind("<Up>", lambda event: r_button_2.invoke()), review_window.bind("<Down>", lambda event: r_button_3.invoke()))
    
    def right(): #和next函数搭配生效
        r_button_1.grid(row=4, column=1)
        r_button_4.grid(row=5, column=1)
        r_label_2.grid_forget()
        r_button_2.grid_forget()
        r_button_3.grid_forget()

        return(review_window.unbind("<Up>"), review_window.unbind("<Down>"), review_window.bind("<Return>", lambda event: r_button_1.invoke()))

    def wrong(): #和next函数搭配生效
        curr_list = today_review_list[0]
        curr_review_index = fmc_recite.word_list[curr_list].curr_review_index
        curr_total_index = curr_list*fmc_recite.num + curr_review_index  #换算成在总表中的索引值

        r_button_1.grid(row=4, column=1)
        r_button_4.grid(row=5, column=1)
        r_label_2.grid_forget()
        r_button_2.grid_forget()
        r_button_3.grid_forget()

        fmc_recite.add_info.iloc[curr_total_index, 0] = fmc_recite.add_info.iat[curr_total_index, 0] + 1 #直接在原表中更新，遗忘次数加1
        return(review_window.unbind("<Up>"), review_window.unbind("<Down>"), review_window.bind("<Return>", lambda event: r_button_1.invoke()))

    def next(): #切换至下一个单词或者单词列表
        curr_list = today_review_list[0]
        curr_review_index = fmc_recite.word_list[curr_list].curr_review_index

        if(fmc_recite.word_list[curr_list].curr_review_index < fmc_recite.word_list[curr_list].word_num-1): #未进入新的word_list
            fmc_recite.list_info.iloc[curr_list, 2] = fmc_recite.list_info.iat[curr_list, 2] + 1  #直接在原表中更新当前背诵到的索引
            fmc_recite.word_list[curr_list].curr_review_index += 1
            curr_review_index += 1
            curr_total_index = curr_list*fmc_recite.num + curr_review_index  #换算成在总表中的索引值

            temp_word_1 = str(fmc_recite.word_list[curr_list].word_table.iat[curr_review_index, 0])
            r_label_1["text"] = temp_word_1
            temp_index_1 = sent_data.loc[sent_data['word'] == temp_word_1].index[0] #确定单词对应的句子的索引
            sysn['text'] = sent_data.iat[temp_index_1, 1]
            sent['text'] = sent_data.iat[temp_index_1, 2]
            text_review.delete(1.0, "end") #先删除原有内容
            text_review.insert(1.0, fmc_recite.add_info.iat[curr_total_index, 1]) #再显示新内容
        else: #进入新的word list，当前表复习完了，这里需要更新word_list对象的repeat_times、finishing_time，同时将刚刚复习完的list从today_review_list中pop出去
            fmc_recite.list_info.iloc[curr_list, 0] = fmc_recite.list_info.iat[curr_list, 0] + 1
            fmc_recite.word_list[curr_list].repeat_times += 1
            fmc_recite.list_info.iloc[curr_list, 2] = 0 #直接在原表中更新当前背诵到的索引，归零！
            
            fmc_recite.list_info.iloc[curr_list, 1] = datetime.date.today().strftime('%Y-%m-%d')
            fmc_recite.word_list[curr_list].finishing_time = datetime.datetime.today()

            today_review_list.pop(0) #将头部元素扔掉
            curr_list = today_review_list[0]
            curr_review_index = fmc_recite.word_list[curr_list].curr_review_index
            curr_total_index = curr_list*fmc_recite.num + curr_review_index  #换算成在总表中的索引值
            r_label_0['text'] = '还没复习的Word List:' + str([i + 1 for i in today_review_list]) #更新显示信息

            temp_word_1 = str(fmc_recite.word_list[curr_list].word_table.iat[curr_review_index, 0])
            r_label_1["text"] = temp_word_1
            temp_index_1 = sent_data.loc[sent_data['word'] == temp_word_1].index[0] #确定单词对应的句子的索引
            sysn['text'] = sent_data.iat[temp_index_1, 1]
            sent['text'] = sent_data.iat[temp_index_1, 2]
            text_review.delete(1.0, "end") #先删除原有内容
            text_review.insert(1.0, fmc_recite.add_info.iat[curr_total_index, 1]) #再显示新内容            

    def exit():
        review_window.destroy()

    r_button_0 = Button(review_window, text="开始复习<Space>", font=("微软雅黑",13), fg="blue",command=start)
    r_button_0.grid(row=3, column=1)
    r_button_1 = Button(review_window, text="检查<Enter>", font=("微软雅黑",13), fg="blue", command=check)
    r_button_2 = Button(review_window, text="记对了 ↑", font=("微软雅黑",13), fg="blue", command=lambda:[right(), next()])
    r_button_3 = Button(review_window, text="记错了 ↓", font=("微软雅黑",13), fg="blue", command=lambda:[wrong(), next()])
    r_button_4 = Button(review_window, text="退出复习模式", font=("微软雅黑",13), fg="red", command=exit)
    
    review_window.bind("<space>", lambda event: r_button_0.invoke())

    def callback():
        pass

    review_window.protocol('WM_DELETE_WINDOW',callback) #禁用关闭按钮
    review_window.resizable(False, False)
    #review_window.grab_set()

def tips():#备注窗口
    tips_window = Toplevel()
    tips_window.title('备注工具') #备注
    tips_window.resizable(False, False)
    tips_window.tk.call('tk', 'scaling', ScaleFactor/75) #设置窗口显示缩放

    text = Text(tips_window, height=6, width=50)
    text.configure(font=("微软雅黑", 12))
    refresh_time = Label(tips_window)
    refresh_time.grid(row=2)

    def read():
        curr_total_index = fmc_recite.curr_list*fmc_recite.num + fmc_recite.curr_word_index - 1  #先更新总表中的索引
        text.delete(1.0, "end") #先删除原有内容
        text.insert(1.0, fmc_recite.add_info.iat[curr_total_index, 1]) #再显示新内容

    def save():
        curr_total_index = fmc_recite.curr_list*fmc_recite.num + fmc_recite.curr_word_index - 1  #先更新总表中的索引
        content = text.get(1.0, "end")
        fmc_recite.add_info.iloc[curr_total_index, 1] = content
        refresh_time['text'] = '上次更新时间：' + str(datetime.datetime.now())[0:19]

    # def auto_save():
    #     curr_total_index = fmc_recite.curr_list*fmc_recite.num + fmc_recite.curr_word_index - 1  #总表中的索引值
    #     text.delete(1.0, "end") #先删除原有内容
    #     text.insert(1.0, fmc_recite.add_info.iat[curr_total_index, 1]) #再显示新内容
    #     content = text.get(1.0, "end")
    #     fmc_recite.add_info.iloc[curr_total_index, 1] = content
    #     refresh_time['text'] = '上次更新时间：' + str(datetime.datetime.now())[0:19]
    #     tips_window.after(1000, auto_save)

    #tips_window.after(1000, auto_save) #每一秒更新一次
    refresh = Button(tips_window, text="刷新", font=("微软雅黑",13), fg="blue", command=read)
    remem = Button(tips_window, text="保存", font=("微软雅黑",13), fg="blue", command=save)
    text.grid(row=0)
    refresh.grid(row=1, padx=(46,160), pady=15) 
    remem.grid(row=1, padx=(160,46), pady=15)

button = Button(my_window, text="开始背诵<Space>", font=("微软雅黑",13),fg="blue", command = fmc_recite.display_word)
button_1 = Button(my_window, text="← 有印象", font=("微软雅黑",13),fg="blue", command = lambda:[remem_button(), button_1.grid_forget(), button_2.grid_forget(), button_3.grid_forget()])
button_2 = Button(my_window, text="无印象 →", font=("微软雅黑",13),fg="blue", command = lambda:[not_remem_button(), button_1.grid_forget(), button_2.grid_forget(), button_3.grid_forget()])
button_3 = Button(my_window, text="保存记录并退出(s)", font=("微软雅黑",13),fg="blue", command = fmc_recite.save_record)
button_4 = Button(my_window, text="直接退出", font=("微软雅黑",13),fg="red", command = sys.exit)
button_5 = Button(my_window, text="今日复习任务(r)", font=("微软雅黑",13),fg="blue", command = open_review)
button_6 = Button(my_window, text="备注", font=("微软雅黑",9), fg="blue", command=tips)
button.grid(row=2)

#设置热键
my_window.bind("<space>", lambda event: button.invoke()) #开始
my_window.bind("<Left>", lambda event: button_1.invoke()) #有印象
my_window.bind("<Right>", lambda event: button_2.invoke()) #无印象
my_window.bind("<Control-s>", lambda event: button_3.invoke()) #保存退出
my_window.bind("<r>", lambda event: button_5.invoke()) #打开复习任务

def callback():
    pass

my_window.protocol('WM_DELETE_WINDOW',callback) #禁用关闭按钮
sentence_window.protocol('WM_DELETE_WINDOW',callback) #禁用关闭按钮
sentence_window.resizable(False, False) #禁止调节大小
my_window.resizable(False, False)
my_window.mainloop() #显示窗口