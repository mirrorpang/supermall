import docx
import os,re

result_path=r'C:\Users\pangyuelong\Desktop\图片'
dir=r'C:\Users\pangyuelong\Desktop\备份'
def read(name,result_path):
    doc = docx.Document(name)
    dict_rel = doc.part._rels
    for rel in dict_rel:
        rel = dict_rel[rel]
        if 'image' in rel.target_ref:
            img_name= re.findall("/(.*)", rel.target_ref)[0]
            word_name = os.path.splitext(name)[0]
            new_name = word_name.split('\\')[-1]
            img_name = f'{new_name}_{img_name}'
            with open(f'{result_path}\\{img_name}', "wb") as f:
                f.write(rel.target_part.blob)

for file in os.listdir(r'C:\Users\pangyuelong\Desktop\备份'):
    file_name = os.path.join(dir, file)
    read(file_name,result_path)

# def get_pictures(word_path, result_path):
#     doc = docx.Document(word_path)
#     dict_rel = doc.part._rels
#     for rel in dict_rel:
#         rel = dict_rel[rel]
#         if "image" in rel.target_ref:
#             if not os.path.exists(result_path):
#                 os.makedirs(result_path)
#             img_name = re.findall("/(.*)", rel.target_ref)[0]
#             word_name = os.path.splitext(word_path)[0]
#             # print(os.sep)
#             if os.sep in word_name:
#                 new_name = word_name.split('\\')[-1]
#             else:
#                 new_name = word_name.split('/')[-1]
#             img_name = f'{new_name}_{img_name}'
#             with open(f'{result_path}/{img_name}', "wb") as f:
#                 f.write(rel.target_part.blob)
