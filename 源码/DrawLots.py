import pygame
from pygame.locals import *
import random
import pandas as pd
import time
import win32api

ButtonDict = {}
df = pd.read_excel(r'./Sources/list.xlsx', sheet_name='list')
data1 = format(df.values)
stu_num = len(df)

def AtContr(x, y):
    #for button in ButtonDict.keys():
    if x in range(255, 345) and y in range(262, 312):
        return button

def AtContrl(x, y):
    #for button in ButtonDict.keys():
    if x in range(20, 90) and y in range(322, 352):
        return lbutton

def AtContrs(x, y):
    #for button in ButtonDict.keys():
    if x in range(20, 90) and y in range(287, 317):
        return sbutton

#初始化pygame
pygame.init()
screen = pygame.display.set_mode((600,375))
# 设置标题
pygame.display.set_caption("抽签")
background = pygame.image.load(r"./Sources/Pics/bg.jpg")
button = pygame.image.load(r"./Sources/Pics/bt.jpg")
button2 = pygame.image.load(r"./Sources/Pics/bt2.jpg")
lbutton = pygame.image.load(r"./Sources/Pics/lbt.jpg")
lbutton2 = pygame.image.load(r"./Sources/Pics/lbt2.jpg")
sbutton = pygame.image.load(r"./Sources/Pics/sbt.jpg")
sbutton2 = pygame.image.load(r"./Sources/Pics/sbt2.jpg")
screen.blit(background, (0, 0))  # 对齐的坐标
but_rect = button.get_rect()  # 获取图片的矩形区域
lbut_rect = lbutton.get_rect()  # 获取图片的矩形区域
sbut_rect = sbutton.get_rect()  # 获取图片的矩形区域
screen_rect = screen.get_rect()  # 获取窗口的矩形区域
but_rect.centerx = screen_rect.centerx  # 将窗口的矩形x坐标值赋值给图片的矩形x坐标值
but_rect.centery = screen_rect.centery+100  # 如上
lbut_rect.centerx = screen_rect.centerx-245  # 将窗口的矩形x坐标值赋值给图片的矩形x坐标值
lbut_rect.centery = screen_rect.centery+150  # 如上
sbut_rect.centerx = screen_rect.centerx-245  # 将窗口的矩形x坐标值赋值给图片的矩形x坐标值
sbut_rect.centery = screen_rect.centery+115  # 如上
screen.blit(button, but_rect)
screen.blit(lbutton, lbut_rect)
screen.blit(sbutton, sbut_rect)
pygame.display.update()  # 显示内容

while True:
    x, y = pygame.mouse.get_pos()
    MyactiveC = AtContr(x, y)
    MyactiveL = AtContrl(x, y)
    MyactiveS = AtContrs(x, y)
    for event in pygame.event.get():
        if event.type == QUIT:
            # 接收到退出事件后退出程序
            exit()
        if MyactiveC:
            if event.type == MOUSEBUTTONDOWN:
                screen.blit(background, (0, 0))
                screen.blit(button, but_rect)
                rand = random.randint(0, stu_num-1)
                font = pygame.font.Font('./Sources/Truetype/t2.TTF', 50)
                text = font.render(df['list'][rand], True, (255, 255, 255), None)
                textRect = text.get_rect()  # 在原位置的基础上改变偏移量
                textRect.centerx = screen.get_rect().centerx  # centerx 是指矩形中心的 X 坐标（就是宽度一半的位置）
                textRect.centery = screen.get_rect().centery-5  # centerx 是指矩形中心的 Y 坐标（就是高度一半的位置）
                screen.blit(text, textRect)
                pygame.display.update()
                screen.blit(button2, but_rect)
                screen.blit(lbutton, lbut_rect)
                screen.blit(sbutton, sbut_rect)
                pygame.display.update()
                time.sleep(0.3)
                screen.blit(button, but_rect)
                screen.blit(lbutton, lbut_rect)
                screen.blit(sbutton, sbut_rect)
                pygame.display.update()
                print(df['list'][rand])
        if MyactiveL:
            if event.type == MOUSEBUTTONDOWN:
                win32api.ShellExecute(0, 'open', '.\Sources\list.xlsx', '', '', 1)
                screen.blit(button, but_rect)
                screen.blit(lbutton2, lbut_rect)
                screen.blit(sbutton, sbut_rect)
                pygame.display.update()
                time.sleep(0.3)
                screen.blit(button, but_rect)
                screen.blit(sbutton, sbut_rect)
                screen.blit(lbutton, lbut_rect)
                pygame.display.update()
        if MyactiveS:
            if event.type == MOUSEBUTTONDOWN:
                win32api.ShellExecute(0, 'open', '.\Sources\Help\index.html', '', '', 1)
                screen.blit(button, but_rect)
                screen.blit(lbutton, lbut_rect)
                screen.blit(sbutton2, sbut_rect)
                pygame.display.update()
                time.sleep(0.3)
                screen.blit(button, but_rect)
                screen.blit(lbutton, lbut_rect)
                screen.blit(sbutton, sbut_rect)
                pygame.display.update()

     # 在screen上绘制image图片，第二个参数为目标位置
    pygame.display.update()
    pygame.display.flip()  # 更新屏幕内容