import fitz
import os
import time
import pandas as pd
import PIL.Image as Image

path = r'C:\Users\pangyuelong\Desktop\源文件\test1.pdf'
doc = fitz.open(path)
t0 = time.clock()
checkXO = r"/Type(?= */XObject)"
checkIM = r"/Subtype(?= */Image)"

imgcount = 0
lenXREF = doc._getXrefLength()

# 打印PDF的信息
print("文件名:{}, 页数: {}, 对象: {}".format(path, len(doc), lenXREF - 1))