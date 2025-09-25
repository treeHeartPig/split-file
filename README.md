该项目用于将txt/word/ppt/xlxs格式文件转为pdf格式，并且保持原文件格式；

/file/upload接口会将上传的文档按页切分，并且每一页会生成缩略图；之后将分的页和缩略图上传到minio中，返回给客户端分页信息和缩略图信息

1、在配置文件中配置minio的信息
2、项目依赖libreoffice，所以必须在本地或服务器上安装libreoffice工具；
3、项目会加载非windows系统下指定的font格式，linux加载的字体文件路径是：/usr/share/fonts/truetype/libreoffice ，该路径可修改，不过修改后里面的字体文件也要存放；
