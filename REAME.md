# word2pdf

脚本说明
--------
这是一个将指定目录下面的所有word文档转换成pdf文件的批量脚本工具

参数说明：
---------
     -s --source    word文档所在目录或者word文档路径
     -s --target    word转换成pdf之后的保存路径，如果不给指定该参数默认保存在与源文件同路径


安装依赖包
--------
pip3 install -r requirements.txt

使用示例：
--------
python3 doc2pdf.py -s d:\学习资料  -t 学习资料pdf