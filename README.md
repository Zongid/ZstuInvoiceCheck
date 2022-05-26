# 浙江理工大学发票查验程序
v1.0.0
……
v2.0.0
1. 合并PDF2PNG、发票查验；
2. 增加异常判断与结束程序；
3. 增加运行错误、完成等messagebox提示；
4. 修改定位元素方式。将Full Xpath改为文本匹配方式；


# 使用说明

## 点击下载->[ZstuInvoiceCheck.zip](https://github.com/Zongid/ZstuInvoiceCheck/releases/download/v2.0.0/ZstuInvoiceCheck_v2.0.0.zip)并解压缩


## 运行配置
### 1. 下载并安装Chrome浏览器

如已安装请忽略此步（[Chrome下载地址](https://www.google.cn/chrome/)）

### 2. 下载浏览器驱动

需根据自己的Chrome浏览器的版本选择下载（链接[Chromedriver](http://npm.taobao.org/mirrors/chromedriver/)）

附件：

* [chromedriver_win32 101.0.4951.15.zip](https://github.com/Zongid/ZstuInvoiceCheck/releases/download/v2.0.0/chromedriver_win32.101.0.4951.15.zip)
* [chromedriver_win32 101.0.4951.41.zip](https://github.com/Zongid/ZstuInvoiceCheck/releases/download/v2.0.0/chromedriver_win32.101.0.4951.41.zip)
* [chromedriver_win32 102.0.5005.27.zip](https://github.com/Zongid/ZstuInvoiceCheck/releases/download/v2.0.0/chromedriver_win32.102.0.5005.27.zip)

查看Chrome版本方式：

![Chromeversion](./img/Chromeversion.png)

### 3. 把chromedriver.exe文件复制到浏览器的安装目录下

例：C:\Program Files (x86)\Google\Chrome\Application    （要根据自己实际安装目录）

### 4. 修改json文件并保存

将  "chromedriverpath":"驱动路径"  中的“驱动路径”修改为第2步中的路径

将  "username":"账号","password":"密码"  中的“账号”、“密码”修改为自己的账号密码

例：{"chromedriverpath":"X:/Google/Chrome/Application/chromedriver.exe","username":"2018327113028","password":"lyzmima"}

## 运行程序

### 5.双击ZstuInvoiceCheck.exe运行
选择发票文件所在文件夹，程序将会把pdf文件转为png图片，然后选择文件夹中的发票图像、提取信息、查验和生成统计结果Excel文件。


# 注意

1. 仅支持浙江理工大学校园网环境；
2. 只需第一次操作1、2、3、4，后续可直接运行程序；
3. 支持**pdf、png、jpg、jpeg、bmp**格式的发票文件，且可混合置于同一文件夹；
4. json文件可用**记事本**、**Sublime Text**等应用程序打开修改；
5. 仅支持**增值税电子普通发票**，深圳区块链发票、机打发票等暂不支持；
6. 图像需包含左上角**二维码**，其他部分可有可无（程序通过识别二维码获取信息）；
7. libiconv.dll、libzbar-64.dll、userdata.json**不可删除或重命名**；
8. 查验结果为<font color=#FF0000>**NULL**</font>（红色加粗）时，说明未能识别出发票信息，需人工查验；
9. 若财务系统再次改版，此程序将不再适用；
10. 其他未尽事宜  <Zongid@outlook.com>
