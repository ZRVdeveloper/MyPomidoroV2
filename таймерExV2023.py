from tkinter import *
import pyexcel as pe
import winsound, math, datetime, customtkinter
from datetime import datetime, date
from random import choice

from customtkinter import *  # <- import the CustomTkinter module

class Timer(customtkinter.CTk):
    
    def on_closing(self):
        if messagebox.askokcancel("Закрити", "Ви хочете закрити выкно"):
            mywin.destroy()
    def write_in_exel_file(self, excel_file, date=None, time = None,size = None, progect = None, comment = None):
    
        sheet = pe.get_sheet(file_name=excel_file)
        sheet.row += [date, time, size, progect, comment]
        sheet.save_as(excel_file)
        pe.get_sheet(file_name=excel_file)
    def open_file(self):
        self.btnPlus.place_forget()
        try:
            f = open("rez1.txt","a")
            print('Open')

        except:
            print('Erorr open file')
            f = open("rez.txt","w")
            print('be problem, but Open')
        finally:
            n = f'\n{datetime.now()} {self.timerS}хв - 1 помідора ({self.comment.get()})'
            f.write(n)
            f.close()
            self.write_in_exel_file ("rezult2023.xls", str(datetime.now()), str(self.timerS)+'хв', 1,str(self.comment.get()), str(self.pro.get()))
        
        self.com.place_forget()
        self.comLab.place_forget()        
        self.proLab.place_forget()
        self.pro.place_forget()
        self.time.place_forget()
        
        
    def timer_pause(self, m = 5):
        if self.pause % 2 == 1:
            self.time.configure(text_color='red')
            self.btnPause.configure(text = "Продовжити", fg_color='red')
            self.after(1000, self.timer_pause)
        
        else:
            self.time.configure(text_color='white')        
            self.btnPause.configure(text = "Пауза", fg_color='blue')        
            self.timer()

    def timer(self, m = 5):
        self.timerStart
        seconds = math.fmod(self.timerStart,60) # Получаем секунды
        minutes = math.fmod(self.timerStart/60, 60) # Получаем минуты
        hour = math.fmod(self.timerStart/60/60, 60) # Получаем часы
    
    
    
        # Условие если время закончилось то...
        if (self.timerStart <= 0):
            # Таймер удаляется
            self.timerStart = 0
            # Выводит сообщение что время закончилось
            print ("Час завершився")
        
            self.btnPlus.place(x=520, y=110)
            self.btn5.place(x=20, y=0)
            self.btn25.place(x=20, y=40)
            self.btncofe.place(x=20, y=80)
            self.proLab.place(x=520, y=10)        
            self.pro.place(x=520, y=30)        
            self.comLab.place(x=520, y=60)
            self.com.place(x=520, y=80)
            self.btnPause.place_forget()
            winsound.PlaySound("clock.wav", winsound.SND_ALIAS)    

        else: # Иначе
            # Создаём строку с выводом времени
            timeNow = str(math.trunc(hour)) + ' : ' + str(math.trunc(minutes)) + ' : ' + str(math.trunc(seconds));
            # Выводим строку в блок для показа таймера
            self.time.configure(text = f'{timeNow}')
            self.timerStart = self.timerStart - 1
            #print (timerStart)
            self.after(1000, self.timer_pause)
        
    def minute(self, m):
        self.btn5.place_forget()
        self.btn25.place_forget()
        self.btncofe.place_forget()
        self.btnPause.place(x=20, y=120)
        self.timerS
        self.timerStart
        self.timerS = m # Кількість хвилин
        self.timerStart = self.timerS * 60 # Кількість хвилин
        self.time.pack()
    
        self.timer(m)

    def pause_plus(self):
        self.pause += 1
    
    def __init__(self):
        
        super().__init__()
        zrv_style = ["blue", "green", "dark-blue"]
        customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
        customtkinter.set_default_color_theme(choice(zrv_style))  # Themes: "blue" (standard), "green", "dark-blue"
        self.pause = 0
        self.timerM = 0
        self.timerS = 0 # Кількість хвилин
        self.timerStart = self.timerS * 60 # Кількість хвилин
        self.geometry("670x150")
        self.btn5 = CTkButton(master = self, text = "5 хвилин", command = lambda: self.minute(5))
        self.btn5.place(x=20, y=0)
        self.btn25 = CTkButton(master = self,text = "25 хвилин", command = lambda: self.minute(25))
        self.btn25.place(x=20, y=40)
        self.btncofe = CTkButton(master = self,text = "кава", command = lambda: self.minute(6))
        self.btncofe.place(x=20, y=80)
        self.btnPlus = CTkButton(master = self,text = "Додати помідор", command = self.open_file)
        self.btnPause = CTkButton(master = self,text = "Пауза", command = self.pause_plus)
        self.project = StringVar()
        self.proLab = CTkLabel (master = self,text = 'Проєкт')
        self.comLab = CTkLabel (master = self,text = 'Робота')
        self.pro = CTkEntry(master = self, textvariable = self.project)
        self.comment = StringVar()
        self.com = CTkEntry(master = self,textvariable = self.comment)
        self.time = CTkLabel(master = self,text = "", font =("Arial", 80))
        
        

if __name__=="__main__":

    timer = Timer()
    
    timer.mainloop()