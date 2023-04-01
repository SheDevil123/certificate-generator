import cv2
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import img2pdf
from PIL import Image


def spacing(s):
    return " "*((32-len(s))//2) +s


FONT_SIZE=1.15
img_path="templates/temp1.jpg"
data_path="sheets/sheet1.xlsx"
output_path="output/"

data = pd.read_excel(data_path)
print(list(data.loc[1]))
columns=data.columns.values
name_coord=(560,440)
event_coord=(460,550)
center_coord=(300,600)
date_coord=(75,650)
year_coord=(125,500)
depart_coord=(550,495)
img=cv2.imread(img_path)
dup=img.copy()
# cv2.imshow("",dup)
# cv2.imshow("",img)
# cv2.waitKey(0)
for i in data.index:
    file_name=data.loc[i,"Name"]
    img=dup.copy()
    cv2.putText(img,data.loc[i,"Year"],year_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    cv2.putText(img," "+spacing(data.loc[i,"Department"]),depart_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    cv2.putText(img," "+spacing(data.loc[i,"Name"]),name_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    cv2.putText(img,"   Crash Course in Python Programming",event_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    cv2.putText(img,"  RIT VIRTUSA Center Of Excellence And TechSparks Club",center_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    cv2.putText(img,"from 9.1.2023 to 13.1.2023",date_coord,cv2.FONT_HERSHEY_DUPLEX,FONT_SIZE,(0,0,0),2)
    pdf_path=output_path+data.loc[i,"Name"]+".jpg"
    cv2.imwrite(pdf_path, img)
    img_path = output_path+data.loc[i,"Name"]+".jpg"
    pdf_path = output_path+"pdf's/"+data.loc[i,"Name"]+".pdf"
    image = Image.open(img_path)
    pdf_bytes = img2pdf.convert(image.filename)
    file = open(pdf_path, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()
    print("Successfully made pdf file")



