#!/usr/bin/python
#-*- coding:utf-8 -*-

import os
import re
import cv2
import sys
import json
import urllib
import datetime
import requests
import numpy as np
from docx import Document
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Ui_MainWindow import Ui_MainWindow
from urllib.request import urlretrieve
from PIL import ImageFont, ImageDraw, Image

class BaiQue(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.InitClassMember()
        self.ConnectWidgetToFun()
    
    def InitClassMember(self):
        self.AppId = "wx38330ee81eefe3da"
        self.AppSecret = "07b9b0ec8cab6746e208295616b80165"
        self.AwardList = []
        self.TitleList = []
        self.AuthorList = []
        self.ReviewsList = []
        self.CommentList = []
        self.ContentList = []
        self.AddressList = []
        self.File = ""
        self.Theme = ""
        self.Digest = ""
        self.Article = ""
        self.FileName = ""
        self.Category = ""
        self.PrintText = ""
        self.AccessToken = ""
        self.PoemNum = 0
        self.CurrentIndex = 0
    
    def ConnectWidgetToFun(self):
        self.pushButton.clicked.connect(self.SelectFile)
        self.pushButton_2.clicked.connect(self.GenOneCertificate)
        self.pushButton_3.clicked.connect(self.GenAllCertificate)
        self.pushButton_4.clicked.connect(self.CheckDocxFormat)
        self.pushButton_5.clicked.connect(self.SelectFile)
        self.pushButton_6.clicked.connect(self.CheckDocxFormat)
        self.pushButton_7.clicked.connect(self.GenSelfMediaArticle)
        self.pushButton_8.clicked.connect(self.ClearOutputWindow)
        self.pushButton_9.clicked.connect(self.ClearOutputWindow)
    
    def StackedWidgetSwitch(self):
        item = self.treeWidget.currentItem()
        if item.text(0) == '一键生成奖状':
            self.stackedWidget.setCurrentIndex(0)
            self.CurrentIndex = 0
        else:
            self.stackedWidget.setCurrentIndex(1)
            self.CurrentIndex = 1
    
    def SelectFile(self):
        self.File = ""
        fname = QFileDialog.getOpenFileName(None, "选择获奖文件", "/", "*.docx")
        if fname[0]:
            try:
                self.File = fname[0]
                if self.CurrentIndex == 0:
                    self.textEdit_3.setText(self.File)
                else:
                    self.textEdit_4.setText(self.File)
            except:
                if self.CurrentIndex == 0:
                    self.textEdit_3.setText("错误！")
                else:
                    self.textEdit_4.setText("错误！")
    
    def getAccessToken(self):
        url = "https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid=%s&secret=%s" % (self.AppId, self.AppSecret)
        result = json.loads(requests.post(url).text)
        if "access_token" in result:
            self.AccessToken = result["access_token"]
            self.textEdit_4.setText("成功获取Access Token！")
            return 1
        else:
            self.textEdit_4.setText(result["errmsg"])
            return 0
    
    def uploadImage(self, ImagePath):
        self.textEdit_4.setText("正在上传图片： " + ImagePath)
        url = "https://api.weixin.qq.com/cgi-bin/material/add_material?access_token=%s&type=image" % self.AccessToken
        files = {'media': open('{}'.format(ImagePath), 'rb')}
        response = requests.post(url, files=files).text
        ImgId = json.loads(response)["media_id"]
        ImgUrl = json.loads(response)["url"]
        return ImgId, ImgUrl
    
    def getImageByName(self, name, count, offset=0):
        url = "https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token=%s" % self.AccessToken
        data = {"type":"image", "offset":offset, "count":count}
        response = requests.post(url, json.dumps(data)).text
        ImgJsonList = json.loads(response)["item"]
        for img in ImgJsonList:
            if img["name"] == name:
                return img["media_id"], img["url"]
    
    def getImageListByCount(self, count, offset=0):
        ImgIdList = []
        ImgUrlList = []
        url = "https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token=%s" % self.AccessToken
        data = {"type":"image", "offset":offset, "count":count}
        response = requests.post(url, json.dumps(data)).text
        ImgJsonList = json.loads(response)["item"]
        for img in ImgJsonList:
            ImgIdList.append(img["media_id"])
            ImgUrlList.append(img["url"])
        return ImgIdList, ImgUrlList
    
    def getImgFromBaidu(self, keyword, num):
        index = 0
        PicPath = []
        PicSize = 100000 # Only select picture which size is more than 100KB
        headers = {
            'Referer': 'https://image.baidu.com/search/index?tn=baiduimage',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36 Edg/83.0.478.50',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
        }
        BaseUrl = 'http://image.baidu.com/search/flip?tn=baiduimage&ipn=r&ct=201326592&cl=2&lm=-1&st=-1&fm=result&fr=&sf=1&fmq=1497491098685_R&pv=&ic=0&nc=1&z=&se=1&showtab=0&fb=0&width=&height=&face=0&istype=2&ie=utf-8&ctd=1497491098685%5E00_1519X735&word=' + keyword
        result = requests.get(BaseUrl,headers=headers).content.decode('utf-8')
        PicUrlList = re.findall('"objURL":"(.*?)",', result, re.S)
        if not os.path.exists('图片'):
            os.mkdir('图片') 
        for url in PicUrlList:
            if index < num:
                PicName = '图片/'+str(index)+'.jpg'
                urlretrieve(url, PicName)
                if os.path.getsize(PicName) > PicSize:
                    PicPath.append(PicName)
                    index += 1
        return PicPath
    
    def GenHtmlFile(self, keyword, TemplateNo):
        PicUrlList = []
        ContentChildList = []
        ReviewsChildList = []
        
        #获取Digest PageFoot ReviewLabel三张图片的url
        #id, DigestUrl = self.uploadImage(Digest)
        id, PageFootUrl = self.uploadImage('template/%s/PageFoot.jpg' % TemplateNo)
        id, ReviewLabelUrl = self.uploadImage('template/%s/Reviews.jpg' % TemplateNo)
        
        #从百度图片中获取若干图片并上传至公众号       
        PicList = self.getImgFromBaidu(keyword, int(self.PoemNum/4))		
        for pic in PicList:
            id, url = self.uploadImage(pic)
            PicUrlList.append(url)
        
        #导入css样式（tbd）
        #f = open('template/%s/style.css' % TemplateNo, 'r')       
        if "诗" in self.Theme:
            Alignment = "center"
        else:
            Alignment = "Left"
        
        #开始生成公众号文章html代码
        #self.Article += f.read() % Alignment
        self.Article += '<section style="padding:10px; background-color:rgb(246, 242, 238)">'
        for i in range(self.PoemNum):
            ContentChildList = re.split(r'-', self.ContentList[i])
            ReviewsChildList = re.split(r'-', self.ReviewsList[i])
            self.Article += '<br><p style="text-align:center; margin-bottom:10px"><span style="font-size:16px; color:rgb(38, 46, 56)"><strong>'
            self.Article += self.TitleList[i]
            self.Article += '</strong></span></p>'
            for j in range(len(ContentChildList)):
                if len(ContentChildList[j]) > 0:    
                    self.Article += '<p style="text-align:%s"><span style="font-size: 16px; color:rgb(38, 46, 56)">' % Alignment
                    self.Article += ContentChildList[j]
                    self.Article += '</span></p>'
            if self.CommentList[i] != "":
                self.Article += '<p><span style="font-size:14px; color:rgb(145, 152, 159)">'
                self.Article += self.CommentList[i]
                self.Article += '</span></p>'
            self.Article += '<img src="%s">' % ReviewLabelUrl
            for k in range(len(ReviewsChildList)):
                if len(ReviewsChildList[k]) > 0:
                    self.Article += '<p><span style="font-size:15px; color:rgb(151, 18, 19)"><strong>'
                    self.Article += ReviewsChildList[k]
                    self.Article += '</strong></span></p>'
            self.Article += '<br>'
            if i%4 == 0 and i != 0:
                self.Article += '<img src="%s"><br>' % PicUrlList[int(i/4)-1]
        self.Article += '<br><img src="%s"></section>' % PageFootUrl
        
        #生成预览html文件
        f = open('preview.html', 'w', encoding='utf-8')
        f.write(self.Article)
        f.close()
    
    def GenAllCertificate(self):
        result = 0
        if self.File == "":
            self.textEdit_3.setText("请先导入文档！")
        else:
            self.ParseDocx3(self.File)
            if len(self.TitleList) == len(self.ReviewsList) == len(self.ContentList):
                self.textEdit_3.setText("文档格式无误，正在自动生成所有奖状！")
                for i in range(len(self.AuthorList)):
                    self.DrawCertificate(self.AwardList[i], self.TitleList[i], self.AuthorList[i], self.Category, self.ContentList[i])
            else:
                self.textEdit_3.setText("文档格式不正确，请选择需要打印的输出信息！")
    
    def GenOneCertificate(self):
        Author = self.lineEdit.text()
        Title = self.lineEdit_2.text()
        Award = self.comboBox.currentText()
        Content = self.textEdit.toPlainText()
        Category = self.comboBox_2.currentText()
        if Content == "" or Author == "" or Title == "":
            self.textEdit_3.setText("请输入必要的作品信息！")
        else:
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
        
        Background = cv2.imread("CertTemplate.jpg")
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
        self.textEdit_3.append("生成：" + Author +".jpg！\n")
    
    def ParseDocx2(self, FileName):
        self.InitClassMember()
        TitleFlag = 0
        Reviews = ""
        Content = ""
        Comment = ""
        RegDigest = "本期评委.*"
        RegTheme = ".*第.*季度.*"
        RegTitle = "[0-9]+、.*|[0-9]+-[1-3]+、.*|[0-9]+·.*|[0-9]+-[1-3]+·.*"
        RegReviews = "..+:.*|..+：.*|【.*"
        RegComment = "^注.*"
        document = Document(FileName)
        for paragraph in document.paragraphs:
            if paragraph.text != "":
                FindDigest = re.search(RegDigest, paragraph.text)
                FindTheme = re.search(RegTheme, paragraph.text)
                FindTitle = re.search(RegTitle, paragraph.text)
                FindReviews = re.search(RegReviews, paragraph.text)
                FindComment = re.search(RegComment, paragraph.text)
                if FindTheme:
                    self.Theme = FindTheme.group()
                elif FindDigest:
                    self.Digest = FindDigest.group()
                elif FindReviews:
                    TitleFlag = 0
                    if Content != "":
                        self.ContentList.append(Content)
                        Content = ""
                    Reviews += FindReviews.group() + "-"
                elif TitleFlag == 1 and FindComment == None:
                    Content += paragraph.text + "-"
                elif TitleFlag == 1 and FindComment:
                    Comment = FindComment.group()
                elif FindTitle:
                    self.TitleList.append(FindTitle.group())
                    TitleFlag = 1
            else:
                if Reviews != "":
                    self.CommentList.append(Comment)
                    self.ReviewsList.append(Reviews)
                    Comment = ""
                    Reviews = ""
        self.PoemNum = len(self.TitleList)
    
    def ParseDocx3(self, File):
        self.InitClassMember()
        Content = ""
        CurrentAward = ""
        AddressFlag = 0
        RegTitle = "[0-9]+、.*|[0-9]+-[1-3]+、.*|[0-9]+·.*|[0-9]+-[1-3]+·.*"
        RegAuthor = "作者：.*|作者:.*|姓名：.*|姓名:.*"
        RegAddress = "地址:|地址：|住址:|住址："
        WritingType = ["诗部", "词部", "曲部"]
        AwardsClass = ["一等奖", "二等奖", "三等奖", "优秀奖"]
        document = Document(File)
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
    
    def UploadArticle(self, cover):
        CoverImgId, url = self.uploadImage(cover)
        FileNames = {
        "articles": [{
            "title": self.Theme,
            "thumb_media_id": CoverImgId,
            "author": "白雀奖",
            "digest": self.Digest + " 本期共有%s首作品入围。" % self.PoemNum,
            "show_cover_pic": 1,
            "content": self.Article,
            "content_source_url": "",
            "need_open_comment": 0,
            "only_fans_can_comment": 0}
        ]}
        url = "https://api.weixin.qq.com/cgi-bin/material/add_news?access_token=%s" % self.AccessToken
        response = requests.post(url, json.dumps(FileNames, ensure_ascii=False).encode('utf-8')).text
        MediaId = json.loads(response)["media_id"]
        self.textEdit_4.setText("文章已成功生成，请进入公众号素材区查看！")
    
    def GenSelfMediaArticle(self):
        result = 0
        if self.File == "":
            self.textEdit_4.setText("请先导入文档！")
        else:
            self.ParseDocx2(self.File)
            if len(self.TitleList) == len(self.ReviewsList) == len(self.ContentList):
                self.textEdit_4.setText("文档格式无误，正在自动生成公众号文章！")
                result = 1
            else:
                self.textEdit_4.setText("文档格式不正确，请选择需要打印的输出信息！")
        if result == 1:
            self.textEdit_4.setText("文档格式不正确，请选择需要打印的输出信息！")
            result = self.getAccessToken()
        if result == 1:
            KeyWord = self.lineEdit_3.text()
            if KeyWord == "":
                self.textEdit_4.setText("请输入图片关键字！")
            else:
                self.GenHtmlFile(KeyWord, 1)
                self.UploadArticle("cover.jpg")
    
    def CheckDocxFormat(self):
        if self.File == "":
            if self.CurrentIndex == 0:
                self.textEdit_3.setText("请先导入文档！")
            else:
                self.textEdit_4.setText("请先导入文档！")
        else:
            if self.CurrentIndex == 0:
                self.ParseDocx3(self.File)
            else:
                self.ParseDocx2(self.File)
        Class = self.comboBox_4.currentText()
        if Class == "标题":
            for item in self.TitleList:
                self.PrintText = self.PrintText + item + "\n" 
                self.textEdit_4.setText(self.PrintText)
        elif Class == "内容":
            for item in self.ContentList:
                self.PrintText = self.PrintText + item + "\n\n" 
                self.textEdit_4.setText(self.PrintText)
        elif Class == "评语":
            for item in self.ReviewsList:
                self.PrintText = self.PrintText + item + "\n\n" 
                self.textEdit_4.setText(self.PrintText)
        elif Class == "注释":
            for item in self.CommentList:
                self.PrintText = self.PrintText + item + "\n" 
                self.textEdit_4.setText(self.PrintText)
    
    def ClearOutputWindow(self):
        self.PrintText = ""
        if self.CurrentIndex == 0:
            self.textEdit_3.setText(self.PrintText)
        else:
            self.textEdit_4.setText(self.PrintText)
    
if __name__ == '__main__':
  app = QApplication(sys.argv)
  bq = BaiQue()
  bq.show()
  sys.exit(app.exec_())