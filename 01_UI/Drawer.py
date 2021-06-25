#!/usr/bin/python
#-*- coding:utf-8 -*-

import re
import cv2
import datetime
import numpy as np
from PIL import ImageFont, ImageDraw, Image

class Drawer():
    def __init__(self,Stamp,Cate,Quar,Year):
        self.Year = str(Year)
        self.Quar = str(Quar)
        self.Cate = str(Cate)
        if Stamp == "带邮戳":
            self.Template = "CertTemplateWithStamp.jpg"
        else:
            self.Template = "CertTemplateWithoutStamp.jpg"
    
    def DrawCertificate(self,Award,Title,Author,Content):
        if self.Cate == "诗部":
            LineList = re.split(r'-',Content.strip())
        else:
            LineList = re.split(r'。',Content[:-2].strip())
        ThisYear = str(datetime.datetime.now().year)[2:]
        Month = str(datetime.datetime.now().month)
        Day = str(datetime.datetime.now().day)
        
        TextFont = ImageFont.truetype("font/华康魏碑.ttc", 60)
        Date1Font = ImageFont.truetype("font/华康宋体.ttc", 70)
        Date2Font = ImageFont.truetype("font/华康宋体.ttc", 50)
        TitleFont = ImageFont.truetype("font/华康黑体.ttc", 70)
        AwardFont = ImageFont.truetype("font/华康楷体.ttc", 140)
        AuthorFont = ImageFont.truetype("font/华康黑体.ttc", 110)
        CateFont = ImageFont.truetype("font/华康魏碑.ttc", 60)
        
        Background = cv2.imread(self.Template)
        PilImage = Image.fromarray(Background)
        draw = ImageDraw.Draw(PilImage)
        
        # 尽量使题目和作者居中
        if len(Title) >= 15:
            draw.text((470, 700), Title, font=TitleFont, fill=(0, 0, 0))
        else:
            draw.text((1000-len(Title)*35, 700), Title, font=TitleFont, fill=(0, 0, 0))
        if len(Author) >= 12:
            draw.text((1800, 700), Author, font=AuthorFont, fill=(0, 0, 0))
        else:
            draw.text((2500-len(Author)*50, 700), Author, font=AuthorFont, fill=(0, 0, 0))
        
        draw.text((810, 940), Author, font=TextFont, fill=(0, 0, 0))
        draw.text((2750, 860), self.Year, font=Date1Font, fill=(0, 0, 0))
        draw.text((2950, 860), self.Quar, font=Date1Font, fill=(0, 0, 0))
        draw.text((2650, 1960), ThisYear, font=Date2Font, fill=(0, 0, 0))
        draw.text((2750, 1960), Month, font=Date2Font, fill=(0, 0, 0))
        draw.text((2850, 1960), Day, font=Date2Font, fill=(0, 0, 0))
        draw.text((2350, 1080), str(Award), font=AwardFont, fill=(0, 0, 0))
        draw.text((2850, 1140), self.Cate, font=CateFont, fill=(0, 0, 0))
        
        for i in range(len(LineList)):
            if len(LineList[i]) > 0:
                if self.Cate == "诗部":
                    draw.text((500, 1050 + 80*i), LineList[i], font=TextFont, fill=(0, 0, 0))
                else:
                    draw.text((500, 1050 + 80*i), LineList[i]+"。", font=TextFont, fill=(0, 0, 0))
        
        Background = np.array(PilImage)  
        cv2.waitKey()
        cv2.imencode(".jpg", Background)[1].tofile(str(Award)+str(Author)+".jpg")