######################################################
# 按顺序自动打开文件夹下的图片，按任意按钮下一张，按e退出程序
######################################################


import os
import cv2
path = 'C:/Users/CSY/Desktop/pic/'
files = os.listdir(path)
for wenjianjia in files:
    pictures = os.listdir(path+wenjianjia)
    print('----------------------------------')
    print(wenjianjia)
    print('----------------------------------')
    for pic in pictures:
        img = cv2.imread(path + wenjianjia + '/' + pic)
        cv2.namedWindow('picture', cv2.WINDOW_NORMAL)
        cv2.resizeWindow('picture', 1280, 720)
        cv2.imshow('picture', img)
        if cv2.waitKey(0) & 0xFF == ord('e'):
            print(pic)
