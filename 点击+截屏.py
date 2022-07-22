# -*- coding: utf-8 -*-
"""
Created on Wed Aug 11 00:01:15 2021

@author: HERO
"""

import pyautogui as pa
import time
  
for i in range(0,5):
    ad0=r'C:\Users\33590\Desktop\222'
    ad=ad0+'\图'+str(i)+'.png'
    pa.click(1773,12)
    time.sleep(2.0)
    print (ad)  
    im=pa.screenshot(ad)
    i=i+1