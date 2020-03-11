# ReplaceWordPic
python批量替换word文档指定位置的图片

1、安装python2.7

2、pip install python-docx

3、pip install lxml


# 思路:
1.word文件重命名为zip文件

2.解压zip到临时文件夹tmp下

3.图片文件（word已经重命名了）在word\media路径下

4.解析word文档脚本，找到对应的id和图片映射关系

5.替换word\media下对应的文件

6.临时文件压缩为zip，再重命名为原word名
