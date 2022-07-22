#coding:utf-8
from PIL import Image
import os

#图片压缩批处理  
def compressImage(srcPath,dstPath):  
    for root, dirs, files in os.walk(srcPath):  
        #print(root)
        for name in files:
            srcFile=os.path.join(root,name)
            #print(root)
            #拼接完整的文件或文件夹路径
            #srcFile=os.path.join(srcPath,filename)
         
            #print(srcFile)
            dstFile=os.path.join(root,name)
            #print (srcFile)
            #print (dstFile)
            #print(dstPath)
    
            #如果是文件就处理
            if srcFile.endswith(".jpg"):     
                #打开原图片缩小后保存，可以用if srcFile.endswith(".jpg")或者split，splitext等函数等针对特定文件压缩
                sImg=Image.open(srcFile)  
                w,h=sImg.size  
                #print (w,h)
                dImg=sImg.resize((w//4,h//4),Image.ANTIALIAS)  #设置压缩尺寸和选项，注意尺寸要用括号
                dImg.save(dstFile) #也可以用srcFile原路径保存,或者更改后缀保存，save这个函数后面可以加压缩编码选项JPEG之类的
                #print (dstFile+" compressed succeeded")
            else:
                print(srcFile+"操作异常")
            

compressImage(r'C:\Users\33590\Desktop\song\aaaaa', r'C:\Users\33590\Desktop\song\1')
print("已经完成全部图片的压缩！")