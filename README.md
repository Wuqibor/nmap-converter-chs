# Nmap-Converter-Chs
源项目: https://github.com/mrschyte/nmap-converter  
参考项目: https://github.com/0xn0ne/nmapReport  

本项目是一个用于将 nmap 的 xml 报告转换为 xlsx 文件的Python脚本  
基于原作者修改而来，添加了中文支持和部分小优化  
参数部分参考了另一个项目的优化方案  

# 依赖问题
```bash 
pip install python-libnmap XlsxWriter
```
或者 
```bash 
pip install -r requirements.txt
```
对于部分中文 Windows10(+) 系统在安装 python-libnmap 库时出现 gbk 编码的读入文件解码问题请参考以下博文执行下载库文件修改再重新打包安装:  
[https://blog.csdn.net/zhangpeterx/article/details/88663052](https://blog.csdn.net/zhangpeterx/article/details/88663052 "https://blog.csdn.net/zhangpeterx/article/details/88663052")

# 用法
```bash
用法: nmap-converter-chs.py [-h] [-o xlsx] -r xml

必要参数:
  xml                   nmap 输出的 xml 文件路径

可选参数:
  -h, --help            查看帮助
  -o xlsx, --output xlsx  输出的 xlsx 文件存放目录
```
