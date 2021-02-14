#!/usr/bin/python
#-*- coding:utf-8 -*-

import re
import cv2
import sys
import datetime
import numpy as np
from docx import Document
from PIL import ImageFont, ImageDraw, Image
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Ui_MainWindow import Ui_MainWindow

class AutoCertificate(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.file = ""
        self.Category = ""
        self.AwardList = []
        self.TitleList = []
        self.AuthorList = []
        self.ContentList = []
        self.PushButton1.clicked.connect(self.GetCertificateInfo)
        self.PushButton2.clicked.connect(self.SelectFile)
        self.PushButton3.clicked.connect(self.GenAllCertificate)
    
    def SelectFile(self):
        self.file = ""
        self.Category = ""
        self.AwardList = []
        self.TitleList = []
        self.AuthorList = []
        self.ContentList = []
        fname = QFileDialog.getOpenFileName(None, "选择获奖文件", "/", "*.docx")
        if fname[0]:
            try:
                self.file = fname[0]
                self.TextEdit2.setText(self.file)
            except:
                self.TextEdit2.setText("错误！")
    
    def GetCertificateInfo(self):
        Author = self.LineEdit1.text()
        Title = self.LineEdit2.text()        
        Award = self.ComboBox2.currentText()
        Content = self.TextEdit1.toPlainText()
        Category = self.ComboBox1.currentText()
        self.DrawCertificate(Award, Title, Author, Category, Content)
    
    def DrawCertificate(self, Award, Title, Author, Category, Content):
        LineList = re.split(r'[。]',Content.strip())
        
        Year = str(datetime.datetime.now().year)[:2]
        Month = datetime.datetime.now().month
        Day = str(datetime.datetime.now().day)
        
        if Month < 4 :
            Quarter = "一"
        elif Month > 3 and Month < 7:
            Quarter = "二"
        elif Month > 6 and Month < 10:
            Quarter = "三"
        else :
            Quarter = "四"
        
        TextFont = ImageFont.truetype("font/华康魏碑.ttc", 60)
        Date1Font = ImageFont.truetype("font/华康宋体.ttc", 70)
        Date2Font = ImageFont.truetype("font/华康宋体.ttc", 50)
        TitleFont = ImageFont.truetype("font/华康黑体.ttc", 70)
        AwardFont = ImageFont.truetype("font/华康楷体.ttc", 140)
        AuthorFont = ImageFont.truetype("font/华康黑体.ttc", 110)
        CategoryFont = ImageFont.truetype("font/华康魏碑.ttc", 60)
        
        Background = cv2.imread("template.jpg")
        PilImage = Image.fromarray(Background)
        draw = ImageDraw.Draw(PilImage)
        
        if len(Title) >= 15:
            draw.text((470, 700), Title, font=TitleFont, fill=(0, 0, 0))
        else:
            draw.text((1000-len(Title)*35, 700), Title, font=TitleFont, fill=(0, 0, 0))
        
        if len(Author) >= 12:
            draw.text((1800, 700), Author, font=AuthorFont, fill=(0, 0, 0))
        else:
            draw.text((2500-len(Author)*50, 700), Author, font=AuthorFont, fill=(0, 0, 0))
        
        draw.text((810, 940), Author, font=TextFont, fill=(0, 0, 0))
        draw.text((2750, 860), Year, font=Date1Font, fill=(0, 0, 0))
        draw.text((2950, 860), Quarter, font=Date1Font, fill=(0, 0, 0))
        draw.text((2650, 1960), Year, font=Date2Font, fill=(0, 0, 0))
        draw.text((2750, 1960), str(Month), font=Date2Font, fill=(0, 0, 0))
        draw.text((2850, 1960), Day, font=Date2Font, fill=(0, 0, 0))
        draw.text((2350, 1080), Award, font=AwardFont, fill=(0, 0, 0))
        draw.text((2850, 1140), Category, font=CategoryFont, fill=(0, 0, 0))
        
        for i in range(len(LineList)):
            if len(LineList[i]) > 0:
                draw.text((500, 1050 + 80*i), LineList[i] + "。", font=TextFont, fill=(0, 0, 0))
        
        Background = np.array(PilImage)  
        cv2.waitKey()
        cv2.imencode(".jpg", Background)[1].tofile(str(Author)+".jpg")
        self.TextEdit2.append("即将生成：" + Author +".jpg！")
    
    def ParseDocx(self, file):
        Content = ""
        CurrentAward = ""
        AddressFlag = 0
        RegTitle = "[0-9]+、.*|[0-9]+-[1-3]+、.*|[0-9]+·.*|[0-9]+-[1-3]+·.*"
        RegAuthor = "作者：.*|作者:.*|姓名：.*|姓名:.*"
        RegAddress = "地址:|地址：|住址:|住址："
        WritingType = ["诗部", "词部", "曲部"]
        AwardsClass = ["一等奖", "二等奖", "三等奖", "优秀奖"]
        document = Document(file)
        for paragraph in document.paragraphs:
            if paragraph.text != "":
                FindTitle = re.search(RegTitle, paragraph.text)
                FindAuthor = re.search(RegAuthor, paragraph.text)
                FindAddress = re.search(RegAddress, paragraph.text)
                if self.Category == "":
                    for i in range(len(WritingType)):
                        FindCategory = re.search(WritingType[i], paragraph.text)
                        if FindCategory:
                            self.Category = WritingType[i]
                            break;
                for j in range(len(AwardsClass)):
                    FindAward = re.search(AwardsClass[j]+".*", paragraph.text)
                    if FindAward:
                        CurrentAward = AwardsClass[j]
                        break;
                if FindTitle:
                    self.TitleList.append(FindTitle.group())
                elif FindAuthor:
                    self.AuthorList.append(FindAuthor.group())
                elif FindAddress:
                    AddressFlag = 1
                elif AddressFlag == 1:
                    Content += paragraph.text
            elif Content != "" and AddressFlag == 1:
                self.AwardList.append(CurrentAward)
                self.ContentList.append(Content)
                AddressFlag = 0
                Content = ""
    
    def GenAllCertificate(self):
        self.ParseDocx(self.file)
        if len(self.AwardList) == len(self.TitleList) == len(self.AuthorList) == len(self.ContentList):
            for i in range(len(self.AuthorList)):
               #self.DrawCertificate(self.AwardList[i], self.TitleList[i], self.AuthorList[i], self.Category, self.ContentList[i])
        else:
            self.TextEdit2.setText("请检查文档！")       

if __name__ == '__main__':
    app = QApplication(sys.argv)  
    ui = AutoCertificate()
    ui.show()
    sys.exit(app.exec_())	