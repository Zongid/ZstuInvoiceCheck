"""
Author:Li Yaozong
2022.05.19          程序合并    Add:messagebox、except
2022.05.20          删除refresh
2022.05.22          删除多余finally     添加len(pdffilelist)>0?
2022.05.23          添加try:driver=……
2022.05.24          调整生成excel文件的位置(置于成功打开浏览器之后)     
                    修改定位元素的方式      
                    将绝对路径修改为placeholder内容、text内容
2022.06.10          添加检查、更新ChromeDriver功能
2022.06.13          更新部分messagebox信息      
                    修改main()->main1()#无界面
                    main2()可视化界面！！！！！
2022.06.14          高度修改为可变      label text自动换行
2022.06.20          定义main2中调用函数     更新driver后修改label
2022.07.06          添加Checkbutton
2022.09.19          添加注释
"""
import json
import os
import re
import shutil
import ssl
import sys
import time
import tkinter.messagebox
import winreg
import zipfile
from tkinter import *
from tkinter import filedialog

import cv2 as cv
import fitz  # fitz就是pip install PyMuPDF
import goto
from matplotlib.pyplot import pink, text
import numpy as np
from pyparsing import White
import requests
import xlsxwriter
from goto import goto, label, with_goto
from PIL import Image
from pyzbar import pyzbar as pyzbar
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


# 检测二维码所在位置
def detect(img):
    """
    return:二维码边界框
    """
    barcodes = pyzbar.decode(img)
    for barcode in barcodes:
        # 提取二维码的边界框的位置
        # 画出图像中条形码的边界框
        (x, y, w, h) = barcode.rect
        # 目标区域y1:y2,x1:x2
        img_dst = img[y - 5 : y + h + 5, x - 5 : x + w + 5]
    return img_dst, x, y, w, h


# 读取图片
def cv_imread(file_path):
    """
    file_path:文件路径
    """
    cv_img = cv.imdecode(np.fromfile(file_path, dtype=np.uint8), -1)
    return cv_img


# 提取二维码并解析信息
def Getinfo(filePath):
    """
    return:二维码中的信息组
    """
    img = cv_imread(filePath)
    gray = cv.cvtColor(img, cv.COLOR_BGR2GRAY)
    img_dst, x, y, w, h = detect(gray)
    cv.rectangle(img, (x - 25, y - 25), (x + w + 25, y + h + 25), (0, 0, 255), 8)
    barcodes = pyzbar.decode(img_dst)
    for barcode in barcodes:
        barcodeData = barcode.data.decode("utf-8")
        return barcodeData.split(",")


# 获取当前时间
def getTime():
    """
    return:当前时间 年-月-日 时:分:秒
    """
    t = time.localtime()
    res = (
        str(t.tm_year)
        + "-"
        + str(t.tm_mon)
        + "-"
        + str(t.tm_mday)
        + " "
        + str(t.tm_hour)
        + ":"
        + str(t.tm_min)
        + ":"
        + str(t.tm_sec)
    )
    return res


# 获取当前ChromeDriver的存放路径
def get_path():
    """
    return: ChromeDriver当前路径
    """
    ChromeDriverLocating = os.popen("where chromedriver").read()
    ChromeSavePath, ChromeName = os.path.split(ChromeDriverLocating)
    return ChromeSavePath


# 查验发票
# 通过Xpath定位元素，健壮性较弱
@with_goto
def check01(data, file_path):
    with open("userdata.json", "r") as f:
        for line in f:
            temp = json.loads(line)
            chromedriverpath = temp["chromedriverpath"]  # 取出特定键的值
            username = temp["username"]
            password = temp["password"]
    # 创建工作薄
    asveFile = "发票查验结果统计" + getTime().replace(":", "-") + ".xlsx"
    workbook = xlsxwriter.Workbook(asveFile)
    # 创建工作表
    worksheet = workbook.add_worksheet()
    worksheet.set_column("A:H", 20)
    worksheet.write(0, 0, "发票代码")
    worksheet.write(0, 1, "发票号码")
    worksheet.write(0, 2, "发票日期")
    worksheet.write(0, 3, "发票金额（不含税）")
    worksheet.write(0, 4, "校验码")
    worksheet.write(0, 5, "查验时间")
    worksheet.write(0, 6, "查验结果")
    worksheet.write(0, 7, "文件路径")

    opt = webdriver.ChromeOptions()  # 创建浏览器
    # opt.set_headless()                            #无窗口模式
    # opt.add_argument('headless')
    s = Service(chromedriverpath)
    driver = webdriver.Chrome(service=s, options=opt)  # 创建浏览器对象
    driver.get("https://i.zstu.edu.cn/browse")  # 打开网页
    driver.maximize_window()  # 最大化窗口

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div/div/div/div[1]/div[3]/div/span")
            )
        )
    finally:
        driver.find_element(
            By.XPATH, "/html/body/div/div/div/div[1]/div[3]/div/span"
        ).click()
        time.sleep(1)

    # 输入账号密码
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-right-root/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[1]/app-login-normal/div/form/div[1]/nz-input-group/input",
                )
            )
        )
    finally:
        driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-right-root/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[1]/app-login-normal/div/form/div[1]/nz-input-group/input",
        ).send_keys(username)

    driver.find_element(
        By.XPATH,
        "/html/body/app-root/app-right-root/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[1]/app-login-normal/div/form/div[2]/nz-input-group/input",
    ).send_keys(password)
    driver.find_element(
        By.XPATH,
        "/html/body/app-root/app-right-root/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[1]/app-login-normal/div/form/div[6]/div/button",
    ).click()
    time.sleep(1)

    # 跳转财务系统
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="root"]/div/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div[4]',
                )
            )
        )
    finally:
        driver.find_element(
            By.XPATH,
            '//*[@id="root"]/div/div/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div[4]',
        ).click()
        # driver.close()

    # 跳转页面
    driver.switch_to.window(driver.window_handles[-1])

    # 点击预约报账
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[3]/a/div/div",
                )
            )
        )
    finally:
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[3]/a/div/div",
        ).click()

    driver.switch_to.frame(1)
    # 点击增值税发票查验
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/table/tbody/tr/td[1]/div[2]/ul/li/ul/li[6]")
            )
        )
    finally:
        driver.find_element(
            By.XPATH, "/html/body/table/tbody/tr/td[1]/div[2]/ul/li/ul/li[6]"
        ).click()

    # 下拉框
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]/select",
                )
            )
        )
    finally:
        select = Select(
            driver.find_element(
                By.XPATH,
                "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[1]/td[2]/select",
            )
        )
        select.select_by_visible_text("普通电子发票或专用发票")

    for i in range(len(data)):
        result = "NULL"
        font = workbook.add_format({"bold": 0, "color": "black"})
        if data[i] == False:
            font = workbook.add_format({"bold": 1, "color": "red"})
            goto.la
        inCode = data[i][2]
        inNum = data[i][3]
        inDate = data[i][5]
        inAmount = data[i][4]
        inCheckcode = data[i][6][-6:]

        # 填写发票信息
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input",
                    )
                )
            )
        finally:
            # 发票代码
            driver.find_element(
                By.XPATH,
                "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input",
            ).clear()
            # driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input").send_keys(Keys.DELETE)
            driver.find_element(
                By.XPATH,
                "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input",
            ).send_keys(inCode)
        # 发票号码
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[3]/td[2]/input",
        ).clear()
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[3]/td[2]/input",
        ).send_keys(inNum)
        # 开票日期(格式：yyyymmdd,如 20170101)
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[4]/td[2]/input",
        ).clear()
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[4]/td[2]/input",
        ).send_keys(inDate)
        # 发票金额(不含税)
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[5]/td[2]/input",
        ).clear()
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[5]/td[2]/input",
        ).send_keys(inAmount)
        # 校验码（输入校验码后六位）
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[6]/td[2]/input",
        ).clear()
        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[6]/td[2]/input",
        ).send_keys(inCheckcode)

        driver.find_element(
            By.XPATH,
            "/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[2]/button[2]/div",
        ).click()

        try:
            wait = WebDriverWait(driver, 10).until(EC.alert_is_present())
            result = driver.switch_to.alert.text
            driver.switch_to.alert.accept()
        finally:
            print(result.replace(" ", ""))
        worksheet.write(i + 1, 0, inCode)
        worksheet.write(i + 1, 1, inNum)
        worksheet.write(i + 1, 2, inDate)
        worksheet.write(i + 1, 3, inAmount)
        worksheet.write(i + 1, 4, data[i][6])
        label.la
        worksheet.write(i + 1, 5, getTime())
        worksheet.write(i + 1, 6, result.replace(" ", ""), font)
        worksheet.write_url(i + 1, 7, file_path[i])
    worksheet.write(len(data) + 3, 7, "@ZSTU SIST LYZ")
    workbook.close()
    driver.quit()


# 发票查验
@with_goto
def check(data, file_path, save_path, label_exitchrome):
    """
    data:发票信息组
    file_path:发票图像路径组
    save_path:保存excel文件路径
    label_exitchrome:查验完成后是否退出
    """
    try:
        # 尝试打开json文件获取用户账号与密码
        with open("userdata.json", "r") as f:
            for line in f:
                temp = json.loads(line)
                username = temp["username"]  # 取出账号username的值
                password = temp["password"]  # 取出密码password的值
    except:
        # 打开json文件失败，显示提示并退出
        tkinter.messagebox.showerror("提示", "未找到userdata.json文件！")
        sys.exit()
    try:
        # 尝试获取chromedriver
        chromedriverpath = get_path() + "/chromedriver.exe"
        print(chromedriverpath)
    except:
        # 获取chromedriver失败，显示提示并退出
        tkinter.messagebox.showerror("提示", "未找到chromedriver！")
        sys.exit()

    opt = webdriver.ChromeOptions()  # 创建浏览器
    # opt.set_headless()                            #无窗口模式
    # opt.add_argument('headless')
    s = Service(chromedriverpath)
    try:
        driver = webdriver.Chrome(service=s, options=opt)  # 创建浏览器对象
    except:
        tkinter.messagebox.showerror("ERROR", "打开浏览器失败！\n请确认是否已下载chromedriver.exe...")
        sys.exit()
    try:
        # 尝试打开浙江理工大学信息门户网站
        driver.get("https://i.zstu.edu.cn/browse")  # 打开网页
    except:
        tkinter.messagebox.showerror("ERROR", "打开网页失败！\n请确认是否已连接校园网...")
        return
    # 最大化窗口
    driver.maximize_window()

    # 尝试通过文本寻找登录按钮
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, ".//span[text()='登录']"))
        )
    except:
        tkinter.messagebox.showerror("ERROR", "未找到登录按钮！")
        sys.exit()
    # 点击登录
    driver.find_element(By.XPATH, ".//span[text()='登录']").click()
    time.sleep(1)

    # 输入账号密码
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    ".//input[@placeholder='请输入学工号']",
                )
            )
        )
    except:
        tkinter.messagebox.showerror("ERROR", "未找到账号输入框！")
        sys.exit()
    # 输入账号
    driver.find_element(
        By.XPATH,
        ".//input[@placeholder='请输入学工号']",
    ).send_keys(username)
    # 输入密码
    driver.find_element(
        By.XPATH,
        ".//input[@placeholder='请输入密码']",
    ).send_keys(password)
    # 点击登录
    driver.find_element(
        By.XPATH,
        ".//span[text()='登录']",
    ).click()
    time.sleep(1)

    # 跳转财务系统
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    ".//span[text()='财务系统']",
                )
            )
        )
    except:
        tkinter.messagebox.showerror("ERROR", "跳转财务系统失败！\n请确认账号、密码是否正确...")
        sys.exit()
    # 通过文本寻找财务系统按钮并点击
    driver.find_element(
        By.XPATH,
        ".//span[text()='财务系统']",
    ).click()

    # 跳转页面
    try:
        driver.switch_to.window(driver.window_handles[-1])
    except:
        tkinter.messagebox.showerror("ERROR", "跳转页面失败！")
        sys.exit()

    # 点击预约报账
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    ".//span[text()='预约报账']",
                )
            )
        )
    except:
        tkinter.messagebox.showerror("ERROR", "未找到预约报账！")
        sys.exit()
    # 点击预约报账
    driver.find_element(
        By.XPATH,
        ".//span[text()='预约报账']",
    ).click()

    driver.switch_to.frame(1)
    # 点击增值税发票查验
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, ".//a[text()='增值税发票查验']"))
        )
    except:
        tkinter.messagebox.showerror("ERROR", "未找到增值税发票查验选项！")
        sys.exit()
    driver.find_element(By.XPATH, ".//a[text()='增值税发票查验']").click()

    # 下拉框选择与赋值
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "formWF_YB_19642_d-invoice_type",
                )
            )
        )
    except:
        tkinter.messagebox.showerror("ERROR", "未找到下拉框！")
        sys.exit()
    # 通过ID寻找发票类型下拉框
    select = Select(
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-invoice_type",
        )
    )
    # select.select_by_visible_text("普通电子发票或专用发票")
    select.select_by_value("PT")

    # 创建工作薄
    asveFile = save_path + "发票查验结果统计" + getTime().replace(":", "-") + ".xlsx"
    workbook = xlsxwriter.Workbook(asveFile)
    # 创建工作表
    worksheet = workbook.add_worksheet()
    # 设置列宽
    worksheet.set_column("A:F", 20)
    worksheet.set_column("G:H", 30)
    # 写表头
    worksheet.write(0, 0, "发票代码")
    worksheet.write(0, 1, "发票号码")
    worksheet.write(0, 2, "发票日期")
    worksheet.write(0, 3, "发票金额（不含税）")
    worksheet.write(0, 4, "校验码")
    worksheet.write(0, 5, "查验时间")
    worksheet.write(0, 6, "查验结果")
    worksheet.write(0, 7, "文件路径")

    # 依次查验每条发票信息
    for i in range(len(data)):
        result = "NULL"
        font = workbook.add_format({"bold": 0, "color": "black"})
        # False表示未解析出发票信息，直接跳过查验
        if data[i] == False:
            font = workbook.add_format({"bold": 1, "color": "red"})
            goto.la
        inCode = data[i][2]
        inNum = data[i][3]
        inDate = data[i][5]
        inAmount = data[i][4]
        # 校验码后六位
        inCheckcode = data[i][6][-6:]

        # 填写发票信息
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "formWF_YB_19642_d-fpdm",
                    )
                )
            )
        except:
            tkinter.messagebox.showerror("ERROR", "未找到输入框！")
            sys.exit()
        # 填写发票代码
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fpdm",
        ).clear()
        # driver.find_element(By.XPATH,"/html/body/table/tbody/tr/td[2]/div/div[2]/div[1]/div[1]/form/div[1]/table/tbody/tr/td/div/table/tbody/tr[2]/td[2]/input").send_keys(Keys.DELETE)
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fpdm",
        ).send_keys(inCode)
        # 填写发票号码
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fphm",
        ).clear()
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fphm",
        ).send_keys(inNum)
        # 开票日期(格式：yyyymmdd,如 20170101)
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-kprq",
        ).clear()
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-kprq",
        ).send_keys(inDate)
        # 填写发票金额(不含税)
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fpje",
        ).clear()
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-fpje",
        ).send_keys(inAmount)
        # 填写校验码（输入校验码后六位）
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-jym",
        ).clear()
        driver.find_element(
            By.ID,
            "formWF_YB_19642_d-jym",
        ).send_keys(inCheckcode)
        # 点击查验按钮
        driver.find_element(
            By.XPATH,
            ".//span[text()='查验']",
        ).click()

        # 获取查验结果弹窗
        try:
            wait = WebDriverWait(driver, 10).until(EC.alert_is_present())
            result = driver.switch_to.alert.text
            driver.switch_to.alert.accept()
        except:
            tkinter.messagebox.showerror("ERROR", "未找到Alert弹窗！")
            sys.exit()
        # 将发票信息与查验信息填写至查验结果统计excel文件
        worksheet.write(i + 1, 0, inCode)
        worksheet.write(i + 1, 1, inNum)
        worksheet.write(i + 1, 2, inDate)
        worksheet.write(i + 1, 3, inAmount)
        worksheet.write(i + 1, 4, data[i][6])
        # 二维码未解析信息出信息时跳转标签
        label.la
        worksheet.write(i + 1, 5, getTime())
        worksheet.write(i + 1, 6, result.replace(" ", ""), font)
        worksheet.write_url(i + 1, 7, file_path[i])
    worksheet.write(len(data) + 3, 7, "@ZSTU SIST LYZ")
    # 写文件完成关闭
    workbook.close()
    if label_exitchrome:
        # 关闭chrome浏览器窗口
        driver.quit()
    # 查验完成提示
    tkinter.messagebox.showinfo("提示", "发票查验完成！\n详情见Excel文件...")


# PDF转图片
def pyMuPDF_fitz(pdfPath, imagePath):
    """
    pdfPath:PDF文件路径
    imagePath:生成的图片保存路径
    """
    # startTime_pdf2img = datetime.datetime.now()  # 开始时间
    pdfDoc = fitz.open(pdfPath)
    for pg in range(1):
        # 仅PDF第一页
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        zoom_x = 2  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(imagePath)  # 将图片写入指定的文件夹内


# 获取文件夹中的pdf文件
def get_pdf_file(file_name):
    """
    file_name:目标文件夹
    return:PDF文件路径组
    """
    pdflist = []
    for parent, dirnames, filenames in os.walk(file_name):
        # print(parent,dirnames,filenames)
        for filename in filenames:
            if filename.lower().endswith((".pdf")):
                path = os.path.join(parent, filename).replace("\\", "/")
                pdflist.append(path)
        return pdflist


# 获取文件夹中的bmp、png、jpg、jpeg图片
def get_img_file(file_name):
    """
    file_name:目标文件夹
    return:图片（bmp、png、jpg、jpeg）文件路径组
    """
    imagelist = []
    for parent, dirnames, filenames in os.walk(file_name):
        for filename in filenames:
            if filename.lower().endswith((".bmp", ".png", ".jpg", ".jpeg")):
                path = os.path.join(parent, filename).replace("\\", "/")
                imagelist.append(path)
        return imagelist


# pdf转图片后转移路径
def getPdfPath(filepath):
    """
    return:pdf移动路径
    """
    idx = []
    for m in re.finditer("/", filepath):
        idx.append(m.end())
    # print(idx)
    return filepath[: idx[-1]] + "PDF/" + filepath[idx[-1] :]


"""忽略SSL证书警告"""
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
ssl._create_default_https_context = ssl._create_unverified_context
"""ChromeDriver仓库淘宝镜像地址"""
# ChromeDriver_depot_url = r'http://npm.taobao.org/mirrors/chromedriver/'
ChromeDriver_depot_url = (
    r"https://registry.npmmirror.com/binary.html?path=chromedriver/"
)
ChromeDriver_base_url = r"https://registry.npmmirror.com/-/binary/chromedriver/"


# 通过注册表的方式获取Google Chrome的版本
def get_Chrome_version():
    """
    return: 本机Chrome的版本号（如:96.0.4664）
    """
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
    version, types = winreg.QueryValueEx(key, "version")
    # print("本机目前的Chrome版本为:", version)
    return version


# 查询系统内的Chromedriver版本
def get_version():
    """
    return: 本机ChromeDriver的版本（如：92.0.4515）
    """
    ChromeDriverVersion = os.popen("chromedriver --version").read()
    # print("本机目前的Chromedriver版本为:", ChromeDriverVersion.split(" ")[1])
    return ChromeDriverVersion.split(" ")[1]


# 获取ChromeDriver版本仓库中的所有版本并写入列表
def get_server_chrome_versions(url):
    """
    url: 淘宝的ChromeDriver仓库地址
    return:  versionList 版本列表
    """
    versionList = []
    rep = requests.get(url, verify=False).text
    rep_list = json.loads(rep)
    for i in range(len(rep_list)):
        version = rep_list[i]["name"]  # 提取版本号
        versionList.append(version[:-1])  # 将所有版本存入列表
    return versionList


# 下载chromedriver压缩包文件
def download_driver(download_url):
    """
    download_url:  ChromeDriver对应版本下载地址
    """
    driverfile = requests.get(download_url, verify=False)
    with open("chromedriver.zip", "wb") as zip_file:  # 保存文件到脚本所在目录
        zip_file.write(driverfile.content)
        print("下载成功")


# 解压Chromedriver压缩包到指定目录
def unzip_driver(path):
    """
    path: 指定解压目录
    """
    f = zipfile.ZipFile("chromedriver.zip", "r")
    for file in f.namelist():
        f.extract(file, path)


# 检测chromedriver版本是否兼容
def check_update_chromedriver():
    chromeVersion = get_Chrome_version()
    chrome_main_version = int(chromeVersion.split(".")[0])  # chrome主版本号
    driverVersion = get_version()
    driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号
    download_url = ""
    if driver_main_version != chrome_main_version:
        tkinter.messagebox.showinfo(
            "Info",
            "chromedriver版本与chrome浏览器不兼容!\nChrome_Version："
            + chromeVersion
            + "\nDriver_Version： "
            + driverVersion
            + "\n即将更新……",
        )
        print("chromedriver版本与chrome浏览器不兼容，更新中>>>")
        versionList = get_server_chrome_versions(ChromeDriver_base_url)
        if chromeVersion in versionList:
            download_url = (
                f"{ChromeDriver_base_url}{chromeVersion}/chromedriver_win32.zip"
            )
        else:
            for version in versionList:
                if version.startswith(str(chrome_main_version)):
                    download_url = (
                        f"{ChromeDriver_base_url}{version}/chromedriver_win32.zip"
                    )
                    break
            if download_url == "":
                print(
                    r"暂无法找到与chrome兼容的chromedriver版本，请在http://npm.taobao.org/mirrors/chromedriver/ 核实。"
                )
                tkinter.messagebox.showerror(
                    "ERROR", "暂无法找到与chrome兼容的chromedriver版本!\n请您自行登录相关网站下载更新……"
                )
                sys.exit()
        # 下载chromedriver压缩包
        download_driver(download_url=download_url)
        # 获取chromedriver地址
        Chrome_Location_path = get_path()
        print("解压地址为:", Chrome_Location_path)
        # 解压缩
        unzip_driver(Chrome_Location_path)
        # 删除压缩包
        os.remove("chromedriver.zip")
        print("更新后的Chromedriver版本为：", get_version())
        # 更新完成提示
        tkinter.messagebox.showinfo(
            "info", "chromedriver更新完成！\n更新后的版本为：" + get_version()
        )
    else:
        print(r"chromedriver版本与chrome浏览器相兼容，无需更新chromedriver版本！")
        # 无需更新提示
        tkinter.messagebox.showinfo("info", "chromedriver版本与chrome浏览器相兼容，无需更新！")


# 强制更新chromedriver
def Forced_update():
    chromeVersion = get_Chrome_version()
    chrome_main_version = int(chromeVersion.split(".")[0])  # chrome主版本号
    driverVersion = get_version()
    driver_main_version = int(driverVersion.split(".")[0])  # chromedriver主版本号
    download_url = ""
    versionList = get_server_chrome_versions(ChromeDriver_base_url)
    if chromeVersion in versionList:
        download_url = f"{ChromeDriver_base_url}{chromeVersion}/chromedriver_win32.zip"
    else:
        for version in versionList:
            if version.startswith(str(chrome_main_version)):
                download_url = (
                    f"{ChromeDriver_base_url}{version}/chromedriver_win32.zip"
                )
                break
        if download_url == "":
            print(
                r"暂无法找到与chrome兼容的chromedriver版本，请在http://npm.taobao.org/mirrors/chromedriver/ 核实。"
            )
            tkinter.messagebox.showerror(
                "ERROR", "暂无法找到与chrome兼容的chromedriver版本!\n请您自行登录相关网站下载更新……"
            )
            sys.exit()
    # 下载chromedriver压缩包
    download_driver(download_url=download_url)
    # 获取chromedriver地址
    Chrome_Location_path = get_path()
    print("解压地址为:", Chrome_Location_path)
    # 解压缩
    unzip_driver(Chrome_Location_path)
    # 删除压缩包
    os.remove("chromedriver.zip")
    print("更新后的Chromedriver版本为：", get_version())
    # 更新完成提示
    tkinter.messagebox.showinfo("info", "chromedriver更新完成！\n更新后的版本为：" + get_version())


# 无界面
def main1():
    root = Tk().withdraw()
    try:
        # 检测更新chromedriver
        check_update_chromedriver()
    except:
        tkinter.messagebox.showerror("ERROR", "检测更新失败！")
    # 获取文件夹
    filepath = filedialog.askdirectory(title="选择发票文件存放的位置！", initialdir=r"")

    if filepath != "":
        # 获取PDF文件列表
        pdfFilelist = get_pdf_file(filepath)
    else:
        # 未选择文件夹
        print("select is null")
        tkinter.messagebox.showwarning("提示", "未选择文件夹！")
        sys.exit()

    if len(pdfFilelist) > 0:
        # 创建PDF文件夹
        if not os.path.exists(filepath + "/PDF"):  # 判断存放PDF的文件夹是否存在
            os.makedirs(filepath + "/PDF")
        for i in range(len(pdfFilelist)):
            # PDF转PNG
            pyMuPDF_fitz(pdfFilelist[i], pdfFilelist[i][:-3] + "png")
            shutil.move(pdfFilelist[i], getPdfPath(pdfFilelist[i]))
    else:
        # 文件夹中没有PDF文件
        print("No Pdf file!")

    # 获取文件夹中的图片文件
    imgfiles = get_img_file(filepath)

    # data:发票信息组
    data = []
    print("@ZSTU SIST LiYaozong")
    if len(imgfiles) > 0:
        for i in range(len(imgfiles)):
            try:
                data.append(Getinfo(imgfiles[i]))
            except:
                # 读取发票信息失败
                data.append(False)
        print(data)
        # print(imgfiles)
        check(data, imgfiles, "", 1)
    else:
        print("img is null")
        tkinter.messagebox.showwarning("提示", "未找到任何图像！")


# 有界面
def main2():
    root = Tk()
    # 标题
    root.title("ZstuInvoiceCheck")

    # btn1 = Button(root, text="选择", font=("黑体", 18), command=InvoiceCheck)
    # btn1.place(relx=0.55, rely=0.15, relwidth=0.3, relheight=0.1)
    sw = root.winfo_screenwidth()  # 得到屏幕宽度
    sh = root.winfo_screenheight()  # 得到屏幕高度
    # 设置窗口大小
    ww = 610
    wh = 450
    # 设置窗口位置
    x = (sw - ww) / 2
    y = (sh - wh) / 2 - 55
    root.geometry("%dx%d+%d+%d" % (ww, wh, x, y))
    root.resizable(width=False, height=True)  # 不可更改窗口大小

    # label文本、字体设置
    # 标题
    lable_title = Label(root, text="ZSTU发票查验程序", font=("黑体", 25))
    # chrome版本
    lable_Chversion = Label(
        root, text="Chrome_version: ", font=("黑体", 15), width=20, anchor="w"
    )
    # chrome版本值
    lable_Chversion_value = Label(
        root, text="Chrome_version_value", font=("宋体", 15), width=32, anchor="w"
    )
    # chromedriver版本
    lable_Drversion = Label(
        root, text="Driver_version:", font=("黑体", 15), width=20, anchor="w"
    )
    # chromedriver版本值
    lable_Drversion_value = Label(
        root, text="Driver_version_value", font=("宋体", 15), width=32, anchor="w"
    )
    lable_Invoicepath = Label(
        root, text="发票文件夹:", font=("黑体", 15), width=20, anchor="w"
    )
    lable_Invoicepath_value = Label(
        root,
        text="Invoice_Path_value",
        font=("宋体", 15),
        width=32,
        anchor="w",
        wraplength=350,
        justify="left",
    )
    lable_Savepath = Label(root, text="统计结果保存至:", font=("黑体", 15), width=20, anchor="w")
    lable_Savepath_value = Label(
        root,
        text="Savepath_value",
        font=("宋体", 15),
        width=32,
        anchor="w",
        wraplength=350,
        justify="left",
    )
    # 获取当前chrome、driver版本并赋值
    lable_Chversion_value.configure(text=get_Chrome_version())
    lable_Drversion_value.configure(text=get_version())
    # 默认查验结果保存至当前目录
    lable_Savepath_value.configure(
        # text=os.path.split(os.path.realpath(__file__))[0].replace("\\", "/")
        text=os.path.dirname(os.path.abspath(sys.executable)).replace("\\", "/")
        # text="asdfgnbfvdscadfhdggnbgfvxdfedrftghyujmnbgvfdczsdsdhtymknbvzs,khmjgnfbdv456789123fghj"
    )
    # label布局设置
    # lable1.place(relx=0.5, rely=0.15)
    lable_title.grid(row=0, column=0, ipady=10, columnspan=2)
    lable_Chversion.grid(row=1, column=0, padx=20, pady=10, sticky=W)
    lable_Chversion_value.grid(row=1, column=1, ipadx=5, ipady=10, sticky=W)
    lable_Drversion.grid(row=2, column=0, padx=20, pady=10, sticky=W)
    lable_Drversion_value.grid(row=2, column=1, ipadx=5, ipady=10, sticky=W)
    lable_Invoicepath.grid(row=3, column=0, padx=20, pady=10, sticky=W)
    lable_Invoicepath_value.grid(row=3, column=1, ipadx=5, ipady=10, sticky=W)
    lable_Savepath.grid(row=4, column=0, padx=20, pady=10, sticky=W)
    lable_Savepath_value.grid(row=4, column=1, ipadx=5, ipady=10, sticky=W)

    # 查验完成后是否退出
    label_exit = IntVar()
    label_exit.set(1)
    chkbtn_excelfile = Checkbutton(
        root, text="查验后关闭Chrome", font=("黑体", 15), variable=label_exit
    )
    chkbtn_excelfile.grid(row=5, column=0)

    # 是否删除PDF文件
    label_delPDF = IntVar()
    label_delPDF.set(1)
    chkbtn_excelfile = Checkbutton(
        root, text="删除PDF文件", font=("黑体", 15), variable=label_delPDF
    )
    chkbtn_excelfile.grid(row=5, column=1)

    # 选择发票文件所在位置
    def sel_invoicePath():
        filepath = filedialog.askdirectory(title="选择发票文件存放的位置！", initialdir=r"")
        if filepath != "":
            # 更新发票文件夹label值
            lable_Invoicepath_value.configure(text=filepath)
        else:
            tkinter.messagebox.showwarning("Warning", "未选择任何文件夹！")

    btn_sel_InvoicePath = Button(
        root, text="选择发票文件夹", font=("黑体", 15), command=sel_invoicePath, border=TRUE
    )
    btn_sel_InvoicePath.grid(row=6, column=1, padx=15, pady=15)

    # 选择统计结果文件保存位置
    def sel_saveResultPath():
        filepath = filedialog.askdirectory(title="选择查验结果保存的位置！", initialdir=r"")
        if filepath != "":
            # 更新查验结果保存位置label值
            lable_Savepath_value.configure(text=filepath)
        else:
            tkinter.messagebox.showwarning("Warning", "未选择任何文件夹！")

    btn_sel_SaveResultPath = Button(
        root, text="保存结果文件夹", font=("黑体", 15), command=sel_saveResultPath, border=TRUE
    )
    btn_sel_SaveResultPath.grid(row=6, column=0, padx=15, pady=15)

    # 强制更新
    def mian2_check_update_chromedriver():
        check_update_chromedriver()
        # 更新label版本信息
        lable_Chversion_value.configure(text=get_Chrome_version())
        lable_Drversion_value.configure(text=get_version())

    # 强制更新按钮
    btn_update = Button(
        root,
        text="检测更新Driver",
        font=("黑体", 15),
        command=mian2_check_update_chromedriver,
        border=TRUE,
    )
    btn_update.grid(row=7, column=0, padx=15, pady=15)

    # 开始查验发票流程
    def start_check():
        if lable_Invoicepath_value["text"].strip() == "Invoice_Path_value":
            # 尚未选择发票文件夹
            tkinter.messagebox.showwarning("Warning", "请先选择发票所在文件夹！")
        else:
            filepath = lable_Invoicepath_value["text"].strip()
            pdfFilelist = get_pdf_file(filepath)
            if len(pdfFilelist) > 0:
                # 创建PDF文件夹
                if not os.path.exists(filepath + "/PDF"):  # 判断存放PDF的文件夹是否存在
                    os.makedirs(filepath + "/PDF")
                for i in range(len(pdfFilelist)):
                    # PDF转PNG
                    pyMuPDF_fitz(pdfFilelist[i], pdfFilelist[i][:-3] + "png")
                    shutil.move(pdfFilelist[i], getPdfPath(pdfFilelist[i]))
                if label_delPDF.get():
                    # 勾选删除发票PDF文件
                    try:
                        # 尝试删除PDF文件夹
                        shutil.rmtree(filepath + "/PDF")
                    except OSError as e:
                        # 删除文件夹失败
                        print("ERROR: %s - %s." % (e.filename, e.strerror))
                        # 删除失败提醒
                        tkinter.messagebox.showerror(
                            "Error", e.filename + "-" + e.strerror + "\n删除PDF失败，请手动删除！"
                        )
            else:
                # 文件夹中没有PDF文件
                print("No Pdf file!")
            imgfiles = get_img_file(filepath)
            # data:发票信息组
            data = []
            print("@ZSTU SIST LiYaozong")
            if len(imgfiles) > 0:
                for i in range(len(imgfiles)):
                    try:
                        data.append(Getinfo(imgfiles[i]))
                    except:
                        # 读取信息失败
                        data.append(False)
                print(data)
                print(imgfiles)
                check(
                    data,
                    imgfiles,
                    lable_Savepath_value["text"].strip() + "/",
                    label_exit.get(),
                )
            else:
                # 没有找到发票图像
                print("img is null")
                tkinter.messagebox.showwarning("提示", "未找到任何图像！")

    btn_Check = Button(
        root,
        text="开始查验",
        font=("黑体", 15),
        command=start_check,
        border=TRUE,
    )
    btn_Check.grid(row=7, column=1, ipadx=15, ipady=10)

    # 强制更新driver
    def mian2_forced_update():
        Forced_update()
        # 更新label版本信息
        lable_Chversion_value.configure(text=get_Chrome_version())
        lable_Drversion_value.configure(text=get_version())

    btn_force_update = Button(
        root, text="强制更新", font=("黑体", 10), command=mian2_forced_update, border=TRUE
    )
    btn_force_update.grid(row=8, column=0)
    # 版权label
    lable_Lyz = Label(
        root, text="©ZSTU SIST LiYaozong", font=("黑体", 8), width=35, anchor="se"
    )
    lable_Lyz.grid(row=8, column=1, pady=10, sticky="e")
    root.mainloop()


if __name__ == "__main__":
    main2()

# main2()
