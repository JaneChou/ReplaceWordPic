# -*- coding: UTF-8 -*-
import zipfile
import os
import shutil
from xml.dom.minidom import parse
import sys
import time

default_encoding = 'utf-8'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

'''
思路:
1.word文件重命名为zip文件
2.解压zip到临时文件夹tmp下
3.图片文件（word已经重命名了）在word\media路径下
4.解析word文档脚本，找到对应的id和图片映射关系
5.替换word\media下对应的文件
6.临时文件压缩为zip，再重命名为原word名
'''

'''
:param path:源文件
:param zip_path:docx重命名为zip
:param tmp_path:中转文件夹
:return:
'''
def unzipWord(path, zip_path, tmp_path):
    # 将docx文件重命名为zip文件
    os.rename(path, zip_path)
    # 进行解压
    f = zipfile.ZipFile(zip_path, 'r')
    # 将图片提取并保存
    for file in f.namelist():
        # print(file)
        # if 'media' in file:
        #print(file, tmp_path)
        f.extract(file, tmp_path)
    # 释放该zip文件
    f.close()

    # 删除zip文件
    os.remove(zip_path)
    time.sleep(2)


# 将docx文件从zip还原为docx
def zip2word(zip_path, path):
    os.rename(zip_path, path)

# 写xml
def writeDocXml(dom_tree, xml_path):
    try:
        with open(xml_path,'w') as fh:
            dom_tree.writexml(fh,encoding='utf-8')
            print('write doc xml OK!')
    except:
        print('write doc xml error!')

# 获取ID-图片映射关系
def getIdImageMapInfo(zip_path, tmp_path):
    id_image_map = {}
    
    image_map_file = '\\word\\_rels\\document.xml.rels'
    domTree = parse(tmp_path + image_map_file)
    rootNode = domTree.documentElement
    # print(rootNode.nodeName)

    # 所有关联关系
    relationships = rootNode.getElementsByTagName("Relationship")
    for relationship in relationships:
        if relationship.hasAttribute("Target") and 'media' in relationship.getAttribute("Target"):
            # print("Target: ", relationship.getAttribute("Target"))
            if relationship.hasAttribute("Id"):
                # print("Id: ", relationship.getAttribute("Id"))
                id_image_map[relationship.getAttribute("Id")] = relationship.getAttribute("Target")
    
    return id_image_map

# 获取ID-名字映射关系
def getIdNameMapInfo(zip_path, tmp_path):
    id_name_map = {}
    image_map_file = '\\word\\document.xml'
    domTree = parse(tmp_path+ image_map_file)
    rootNode = domTree.documentElement
    # print(rootNode.nodeName)

    # 所有单元格
    cells = rootNode.getElementsByTagName("w:p")
    for cell in cells:
        # 单元格内行数
        cellrs = cell.getElementsByTagName("w:r")

        for cellr in cellrs:    
            pics = cellr.getElementsByTagName("w:drawing")
            if len(pics) > 0:
                # print(pics[0].getElementsByTagName("wp:docPr")[0].getAttribute("descr"))
                valid = False
                for cellr in cellrs:
                    r = cellr.getElementsByTagName("w:t")
                    if len(r) > 0:
                        if '照片' in r[0].childNodes[0].data:
                            valid = True
                            # print(r[0].childNodes[0].data)
                
                if valid == True:
                    ele = pics[0].getElementsByTagName("wp:docPr")[0]
                    if ele.hasAttribute('descr'):
                        ppath = ele.getAttribute("descr")
                        if len(ppath) > 0:
                            id=(ppath.split('IMG_')[1]).split('.JPG')[0]
                            if id in ppath:
                                new_path = ppath.replace(id, str(to_int(id)+1))
                                (pics[0].getElementsByTagName("wp:docPr")[0]).setAttribute('descr', new_path)

                                id = pics[0].getElementsByTagName("a:blip")[0].getAttribute("r:embed")
                                id_name_map[id] = ppath
    writeDocXml(domTree ,tmp_path+ image_map_file)
    return id_name_map        

def to_int(str):
    try:
        int(str)
        return int(str)
    except ValueError: #报类型错误，说明不是整型的
        try:
            float(str) #用这个来验证，是不是浮点字符串
            return int(float(str))
        except ValueError:  #如果报错，说明即不是浮点，也不是int字符串。   是一个真正的字符串
            return False

# 获取替换图片列表
def getReplacePicList(id_image_info, id_name_info, tmp_path):
    replace_image={}
    for name_id in id_name_info:
        for image_id in id_image_info:
            if image_id == name_id:
                #print(name_id, id_name_info[name_id], id_image_info[image_id])
                id=(id_name_info[name_id].split('IMG_')[1]).split('.JPG')[0]
                print('replace: '+ str(id) + ' --> '+ str(to_int(id)+1))
                # 图片文件Id+1作为key
                replace_image[to_int(id)+1] = tmp_path + "\\word\\" +id_image_info[image_id]
            
    return replace_image

# 获取路径下jpg文件列表
def listdir(path):
    image_list=[]  
    for root, dirs, files in os.walk(path):
        for file in files:  
            if os.path.splitext(file)[1] == '.JPG' or  os.path.splitext(file)[1]=='.jpg' or  os.path.splitext(file)[1] == '.jpeg':  
                image_list.append(os.path.join(root, file))  
    return image_list

# 替换image文件
def replaceImageFile(replace_list, path):
    if len(replace_list) < 1:
        return

    image_list = listdir(path)
    for id in replace_list:
        for image_name in image_list:
            if str(id) in image_name:
                # print image_name, replace_list[id] 
                shutil.copy(image_name, replace_list[id])

def  del_file(path):
    filelist=os.listdir(path)                #列出该目录下的所有文件名
    for f in filelist:
        file = os.path.join(path, f)        #将文件名映射成绝对路路径
        if os.path.isfile(file):            #判断该文件是否为文件或者文件夹
            os.remove(file)                 #若为文件，则直接删除
        elif os.path.isdir(file):
            shutil.rmtree(file,True)        #若为文件夹，则删除该文件夹及文件夹内所有文件
    shutil.rmtree(path,True)                 #最后删除img总文件夹

# 压缩tmp文件夹为zip
def zipDir(tmp_path, zip_path, doc_path):
    zip = zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED)
    for path,dirs,files in os.walk(tmp_path):
        # 去掉目标跟路径，只对目标文件夹下边的文件及文件夹进行压缩
        fpath = path.replace(tmp_path,'')

        for filename in files:
            zip.write(os.path.join(path,filename),os.path.join(fpath,filename))
    zip.close()

    zip2word(zip_path, doc_path)

    del_file(tmp_path)
    time.sleep(2)

if __name__ == '__main__':
    # 源文件
    path = r'1.docx'
    # docx重命名为zip
    zip_path = r'1.zip'
    # 中转文件夹
    tmp_path = r'tmp'
    # 图片路径
    pic_path = r'pic'

    # 重命名并解压文件
    unzipWord(path, zip_path, tmp_path)

    id_image_info = getIdImageMapInfo(zip_path, tmp_path)

    # 解析映射文件，获取Id和图片文件映射关系
    id_name_info = getIdNameMapInfo(zip_path, tmp_path)
    # for id in id_name_info:
    #     print(id, id_name_info[id])

    # 解析word内容，获取Id和图片文件映射关系
    replace_list = getReplacePicList(id_image_info, id_name_info, tmp_path)

    print('replace images...')
    # 替换图片文件
    replaceImageFile(replace_list, pic_path)

    print('generate file...')
    
    # 压缩文件并转为docx文件
    zipDir(tmp_path, zip_path, path)
    print('replace success !!!')

