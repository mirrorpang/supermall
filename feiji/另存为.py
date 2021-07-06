from win32com import client as wc
import os
word = wc.Dispatch('Word.Application')
# doc = word.Documents.Open('/FilePath/test.docx')
path=r'C:\Users\pangyuelong\Desktop\原始文件'
path_1=r'C:\Users\pangyuelong\Desktop\备份'
# doc=word.Documents.Open(path+'\\500kV_川泰Ⅱ线_30-31.docx')
# doc.SaveAs(path_1+'\\500kV_川泰Ⅱ线_30-31.docx')
for files,x,name in os.walk(path):
    for i in range(len(name)):
        doc=word.Documents.Open(os.path.join(path,name[i]))
        word.Visible = False
        doc.SaveAs(os.path.join(path_1,name[i] ))
        doc.Close()
# for files in os.listdir(path):
#     dir=path+'\\'+files
#     doc=word.Documents.Open(dir)d
word.Quit()

