# 浙江理工大学发票查验程序

# 使用说明

## 下载[InvoiceCheck.zip](https://github.com/Zongid/InvoiceCheck/archive/refs/heads/main.zip)并解压缩

可仅保留其中的program文件夹，Chromedriver文件夹、img文件夹、README.md可删除

### 1. 下载浏览器驱动

需根据自己的Chrome浏览器的版本选择下载（链接[Chromedriver](http://npm.taobao.org/mirrors/chromedriver/)）

附件：

* [chromedriver_win32 101.0.4951.15.zip](./Chromedriver/chromedriver_win32%20101.0.4951.15.zip)
* [chromedriver_win32 101.0.4951.41.zip](./Chromedriver/chromedriver_win32%20101.0.4951.41.zip)
* [chromedriver_win32 102.0.5005.27.zip](./Chromedriver/chromedriver_win32%20102.0.5005.27.zip)

查看Chrome版本方式：

![Chromeversion](E:/李耀宗/大四/else/发票识别与校验程序/ZstuInvoiceCheck/img/Chromeversion.png)

### 2. 把chromedriver.exe文件复制到浏览器的安装目录下

例：C:\Program Files (x86)\Google\Chrome\Application    （要根据自己实际安装目录）

### 3. 修改json文件并保存

将  "chromedriverpath":"驱动路径"  中的“驱动路径”修改为第2步中的路径

将  "username":"账号","password":"密码"  中的“账号”、“密码”修改为自己的账号密码

例：{"chromedriverpath":"X:/Google/Chrome/Application/chromedriver.exe","username":"2018327113028","password":"lyzmima"}

### 4. 双击InvoiceCheck.exe运行



# 注意

1. 仅支持浙江理工大学校园网环境；
2. 只需第一次操作1、2、3，后续可直接运行程序；
3. 每次可选择**多张**发票图像；
4. json文件可用**记事本**、**Sublime Text**等应用程序打开修改；
5. 仅支持**增值税电子普通发票**，深圳区块链发票、机打发票等暂不支持；
6. 图像需包含左上角**二维码**，其他部分可有可无（程序通过识别二维码获取信息）；
7. libiconv.dll、libzbar-64.dll、userdata.json**不可删除或重命名**；
8. 查验结果为<font color=#FF0000>**NULL**</font>（红色加粗）时，说明未能识别出发票信息，需人工查验；
9. 若财务系统再次改版，此程序将不再适用；
10. 其他未尽事宜  <Zongid@outlook.com>
