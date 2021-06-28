#!/usr/bin/python
#-*- coding:utf-8 -*-

import os
import re
import json
import urllib
import requests
from PyQt5.QtCore import *
from urllib.request import urlretrieve

url_list = [
"https://api.weixin.qq.com/cgi-bin/material/add_news?access_token=%s",
"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid=%s&secret=%s",
"https://api.weixin.qq.com/cgi-bin/material/add_material?access_token=%s&type=image",
"https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token=%s",
'https://image.baidu.com/search/index?tn=baiduimage',
'http://image.baidu.com/search/flip?tn=baiduimage&ipn=r&ct=201326592&cl=2&lm=-1&st=-1&fm=result&fr=&sf=1&fmq=1497491098685_R&pv=&ic=0&nc=1&z=&se=1&showtab=0&fb=0&width=&height=&face=0&istype=2&ie=utf-8&ctd=1497491098685%5E00_1519X735&word='
]

class Thread_1(QThread):
    signal = pyqtSignal(str)
    
    def __init__(self,CoverImgId,title,digest,content,AccessToken):
        super(Thread_1, self).__init__()
        self.CoverImgId = CoverImgId
        self.title = title
        self.digest = digest
        self.content = content
        self.AccessToken = AccessToken
    
    def run(self):
        FileNames = {
        "articles": [{
            "title": self.title,
            "thumb_media_id": self.CoverImgId,
            "author": "白雀奖",
            "digest": self.digest,
            "show_cover_pic": 1,
            "content": self.content,
            "content_source_url": "",
            "need_open_comment": 0,
            "only_fans_can_comment": 0}
        ]}
        url = url_list[0] % self.AccessToken
        response = requests.post(url, json.dumps(FileNames, ensure_ascii=False).encode('utf-8')).text
        try:
          MediaId = json.loads(response)["media_id"]
          
          # 生成成功，将结果返回给主调函数：GzhHandler.UploadFinish
          self.signal.emit("文章已生成，请到公众号素材区查看！")
        except:
          self.signal.emit(str(json.loads(response)))

class Thread_2(QThread):
    signal = pyqtSignal(str)
    
    def __init__(self,Keyword,Cate,Template,TitleList,ContentList,ReviewsList,CommentList,AuthorList,AddressList,Profile):
        super(Thread_2, self).__init__()
        self.Cate = Cate
        self.Keyword = Keyword
        self.Profile = Profile
        self.Template = Template
        self.TitleList = TitleList
        self.CommentList = CommentList
        self.ContentList = ContentList
        self.ReviewsList = ReviewsList
        self.AuthorList = AuthorList
        self.AddressList = AddressList
        AppId = "wx38330ee81eefe3da"
        AppSecret = "07b9b0ec8cab6746e208295616b80165"
        self.gzh = GzhHandler(AppId, AppSecret)
    
    def run(self):
        result = self.gzh.getAccessToken()
        if result == "True":
          self.GenHtmlFile()
        else:
          self.signal.emit("获取AccessToken失败！")
    
    def GenHtmlFile(self):
        PicUrlList = []   # 插图Url列表
        AwardUrlList = [] # 奖次Url列表
        ContentChildList = []
        ReviewsChildList = []
        
        # 上传页眉、页脚图片
        if self.Template == "月度入围":
            id, PageFootUrl = self.gzh.uploadImage("template/2/PageFoot.jpg")
            id, ReviewLabelUrl = self.gzh.uploadImage("template/2/Reviews.jpg")
        elif self.Template == "个人专辑":
            id, Chapter1Url = self.gzh.uploadImage("template/1/Chapter1.png")
            id, Chapter2Url = self.gzh.uploadImage("template/1/Chapter2.png")
            id, PageFootUrl = self.gzh.uploadImage("template/1/PageFoot.jpg")
        elif self.Template == "季度获奖":
            index = 0 # 奖次图片列表下标
            AwardPosition = [1,2,4,7]
            id, ReviewLabelUrl = self.gzh.uploadImage("template/3/Reviews.jpg")
            id, PageFootUrl = self.gzh.uploadImage("template/3/PageFoot.jpg")
            for i in range(4):
                name = 'template/3/Award'+str(i+1)+'.png'
                id, AwardUrl = self.gzh.uploadImage(name)
                AwardUrlList.append(AwardUrl)
        
        # 诗居中对齐，词和曲左对齐
        if self.Cate == "诗":
            Alignment = "center"
        elif self.Cate == "词" or self.Cate == "曲":
            Alignment = "Left"
        
        # 从百度图片中获取若干图片并上传至公众号   
        if self.Template != "季度获奖" and self.Keyword != "":
            PicList = self.gzh.getImgFromBaidu(self.KeyWord, int(len(self.TitleList)/4))		
            for pic in PicList:
                id, url = self.gzh.uploadImage(pic)
                PicUrlList.append(url)
            
        # 设置背景颜色，个人专辑需要插入人物简介
        if self.Template == "月度入围":
            Article = '<section style="padding:10px; background-color:rgb(246, 242, 238)">\n'     
        elif self.Template == "季度获奖":
            Article = '<section style="padding:10px">\n'
        elif self.Template == "个人专辑":
            Article = '<section style="padding:10px; background-color:rgb(234, 234, 234)">'
            Article += '<img src="%s">\n' % Chapter1Url
            Article += '<p style="text-align:Left"><span style="font-size: 15px;">'
            Article += self.Profile
            Article += '</span></p><br>\n'
            Article += '<img src="%s">\n' % Chapter2Url
        
        for i in range(len(self.TitleList)):
            
            # 将正文和评语拆分成列表
            ContentChildList = re.split(r'-', self.ContentList[i])
            if self.Template != "个人专辑":
                ReviewsChildList = re.split(r'-', self.ReviewsList[i])
            
            # 季度获奖需要插入奖次图片
            if self.Template == "季度获奖":
                if index < 4 and (i+1) == AwardPosition[index]:
                    Article += '<img src="%s">\n' % AwardUrlList[index]
                    index += 1
            
            # 排版标题
            Article += '<br><p style="text-align:center; margin-bottom:10px"><span style="font-size:16px; color:rgb(38, 46, 56)"><strong>'
            Article += self.TitleList[i]
            Article += '</strong></span></p>\n'
            
            # 季度获奖需要增加作者和地址信息
            if self.Template == "季度获奖":
                Article += '<p style="text-align:center; margin-bottom:10px"><span style="font-size: 15px; color:rgb(38, 46, 56)">'
                Article += self.AuthorList[i] + '  ' + self.AddressList[i]
                Article += '</span></p>\n'
            
            # 排版正文
            for j in range(len(ContentChildList)):
                if len(ContentChildList[j]) > 0:    
                    Article += '<p style="text-align:%s"><span style="font-size: 16px; color:rgb(38, 46, 56)">' % Alignment
                    Article += ContentChildList[j]
                    Article += '</span></p>\n'
            
            # 排版注释
            if self.CommentList[i] != "":
                Article += '<p><span style="font-size:14px; color:rgb(145, 152, 159)">'
                Article += self.CommentList[i]
                Article += '</span></p>\n'
            
            # 排版评语，个人专辑没有评委点评
            if self.Template != "个人专辑":
                Article += '<img src="%s">' % ReviewLabelUrl
                for k in range(len(ReviewsChildList)):
                    if len(ReviewsChildList[k]) > 0:
                        Article += '<p><span style="font-size:15px; color:rgb(151, 18, 19)"><strong>'
                        Article += ReviewsChildList[k]
                        Article += '</strong></span></p>\n'
                Article += '<br>'
            
            # 每隔四篇作品插入一张图片，季度获奖不需要插入图片
            if self.Template != "季度获奖" and self.Keyword != "":
                if i%4 == 0 and i != 0:
                    Article += '<br><img src="%s"><br>\n' % PicUrlList[int(i/4)-1]
        
        # 插入公众号二维码图片
        Article += '<br><img src="%s"></section>' % PageFootUrl
        
        # 生成预览html文件
        f = open('preview.html', 'w', encoding='utf-8')
        f.write(Article)
        f.close()
        
        # 将Article返回给主调函数：GzhHandler.PreviewFinish
        self.signal.emit(Article)

class GzhHandler(QObject):
    signal_1 = pyqtSignal(str)
    signal_2 = pyqtSignal(str)
    
    def __init__(self,AppId,AppSecret):
        super(GzhHandler, self).__init__()
        self.AppId = AppId
        self.AppSecret = AppSecret
    
    # 接受thread_2的返回结果，再弹射给BaiQue.PreviewFinish
    @pyqtSlot(str)
    def PreviewFinish(self, Article):
        self.signal_1.emit(Article)
    
    # 接受thread_1的返回结果，再弹射给BaiQue.UploadFinish
    @pyqtSlot(str)
    def UploadFinish(self, Result):
        self.signal_2.emit(Result)
    
    def getAccessToken(self):
        url = url_list[1] % (self.AppId, self.AppSecret)
        result = json.loads(requests.post(url).text)
        if "access_token" in result:
            self.AccessToken = result["access_token"]
            return "True"
        else:
            return result["errmsg"]
    
    # 功能：上传图片至公众号的素材区
    # 入口参数：图片路径
    # 返回值：图片的Id和Url
    def uploadImage(self, ImagePath):
        url = url_list[2] % self.AccessToken
        files = {'media': open('{}'.format(ImagePath), 'rb')}
        response = requests.post(url, files=files).text
        ImgId = json.loads(response)["media_id"]
        ImgUrl = json.loads(response)["url"]
        return ImgId, ImgUrl
    
    # 功能：从公众号的素材区按名称获取一张图片
    # 入口参数：图片名称、图片数量、偏移量
    # 返回值：图片的Id和Url
    def getImageByName(self, name, count, offset=0):
        url = url_list[3] % self.AccessToken
        data = {"type":"image", "offset":offset, "count":count}
        response = requests.post(url, json.dumps(data)).text
        ImgJsonList = json.loads(response)["item"]
        for img in ImgJsonList:
            if img["name"] == name:
                return img["media_id"], img["url"]
    
    # 功能：从公众号的素材区按偏移位置获取若干图片
    # 入口参数：图片数量、偏移量
    # 返回值：图片的Id列表和Url列表
    def getImageListByCount(self, count, offset=0):
        ImgIdList = []
        ImgUrlList = []
        url = url_list[3] % self.AccessToken
        data = {"type":"image", "offset":offset, "count":count}
        response = requests.post(url, json.dumps(data)).text
        ImgJsonList = json.loads(response)["item"]
        for img in ImgJsonList:
            ImgIdList.append(img["media_id"])
            ImgUrlList.append(img["url"])
        return ImgIdList, ImgUrlList
    
    # 功能：开启线程1，上传文章至公众号
    # 入口参数：题目、封面图片、摘要、内容
    # 返回值：无
    def UploadArticle(self, title, cover, digest, content):
        CoverImgId, url = self.uploadImage(cover)
        self.thread_1 = Thread_1(CoverImgId,title,digest,content,self.AccessToken)
        self.thread_1.signal.connect(self.UploadFinish)
        self.thread_1.start()
    
    # 功能：开启线程2，生成html预览文件
    # 入口参数：图片关键字、模板、题目列表、内容列表、评语列表、注释列表、诗/词/曲
    # 返回值：无
    def GenHtmlFile(self,KeyWord,Category,Template,TitleList,ContentList,ReviewsList,CommentList,AuthorList,AddressList,Profile):
        self.thread_2 = Thread_2(KeyWord,Category,Template,TitleList,ContentList,ReviewsList,CommentList,AuthorList,AddressList,Profile)
        self.thread_2.signal.connect(self.PreviewFinish)
        self.thread_2.start()
    
    # 功能：从百度图片上爬取图片，只选择大小大于100KB的图片
    # 入口参数：关键字，图片数量
    # 返回值：图片路径列表
    def getImgFromBaidu(self, keyword, num):
        index = 0 # 每张图片用数字命名
        PicPath = []
        PicSize = 100000
        headers = {
            'Referer': url_list[4],
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36 Edg/83.0.478.50',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
        }
        BaseUrl = url_list[5] + keyword
        result = requests.get(BaseUrl,headers=headers).content.decode('utf-8')
        PicUrlList = re.findall('"objURL":"(.*?)",', result, re.S)
        
        # 爬取的图片都放在 "根目录/图片" 文件夹中
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