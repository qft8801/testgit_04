
# -*- coding: UTF-8 -*-
import win32com.client
import win32com.client.dynamic
import os
#我的示例(Template)文档名为 BugCurve.pptx
def PowerPoint():
  ppt = os.path.join(os.getcwd(), "BugCurve.pptx")
  App = win32com.client.Dispatch("PowerPoint.Application")
  App.Visible = True
  Presentation = App.Presentations.Open(ppt)
  mySlide = Presentation.Slides.Add(2, 12)
  #这份 Presentation 要增加一张 Slide, 位置就插入在第二页，
  #也就是封面(第一页)之后， 12 这个神奇的数字就是..
  #告诉PPT 那一页是空白的页面
  img = os.path.join(os.getcwd(), "This_is_Picture.png")
  shape = mySlide.Shapes.AddPicture(img,LinkToFile=False,SaveWithDocument=True,Left=40,Top=100,Width=650,Height=400)
  #mySlide 中要增加一个 框框(shape)指定那个框框的大小和位置, 
  #然后那个 shape 內容要放入图形