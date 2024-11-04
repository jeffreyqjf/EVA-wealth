import pandas as pd
import requests
from PIL import Image
import cv2 as cv
import os
import numpy as np

def download(url, pic_path_name):
    format_file = url.split(".")[-1]
    if format_file in ["jpg", "jpeg", "JPG", "JPEG"]:
        format_file = "jpg"
    #  print(format_file)
        result = requests.get(url)
        img = pic_path_name + "." + format_file
        with open(img, "wb") as f:
            f.write(result.content)
        return
    elif format_file in ["png", "PNG"]:
        format_file = "png"
        #  print(format_file)
        #  先下载为jpg格式
        result = requests.get(url)
        img = pic_path_name + "." + format_file
        with open(img, "wb") as f:
            f.write(result.content)

        #  转换jpg为png
        img_convert = pic_path_name + ".jpg"
        im = Image.open(img)
        im = im.convert('RGB')
        im.save(img_convert, quality=95)
        print("转换了一张png")

        # 删除原来的png
        if os.path.exists(img):
            os.remove(img)
            print("删除了原来的png")

    elif format_file in ["pdf", "PDF"]:
        pass
    else:
        print("文件格式出现错误，请联系程序开发者！")









if __name__ == "__main__":
    excel_path = input("请输入excel的绝对或相对路径:")
    # df = pd.read_excel(excel_path)
    df = pd.read_excel("./286388460_按文本_发表收集1019_46_46.xlsx")
    #  print(df["1、您的姓名："])
    #  print(df["2、订单类型："])
    #  print(df["3、购买日期"])
    #  print(df["4、金额"])
    #  print(df["5、物品"])
    len_of_df = len(df)
    #  print(len_of_df)
    for one in range(len_of_df):
        #  2024.9.3 【56.94】 木板 陈朵
        path_name = df["3、购买日期"][one] + " 【" + str(df["4、金额"][one]) + "】 " + df["5、物品"][one] + " " + df["1、您的姓名："][one]
        if not os.path.exists(path_name):
            os.mkdir(path_name)
        else:
            pass
        #  print(df["6、购买记录截图："])
        #  print(df["7、支付记录截图："])
        #  print(df["9、发票图片："])
        #  print(df["6、购买记录截图："][3])
        """
        处理购买记录截图
        """
        buy_prt_sc = df["6、购买记录截图："][one]
        if buy_prt_sc != "(跳过)":
            download(buy_prt_sc, "./" + path_name + "/购买记录截图")

        """
        处理支付记录截图
        """
        pay_prt_sc = df["7、支付记录截图："][one]
        if pay_prt_sc != "(跳过)":
            download(pay_prt_sc, "./" + path_name + "/支付记录截图")

        """
        处理发票图片
        """
        pdf = df["9、发票图片："][one]
        if pdf != "(空)":
            download(pdf, "./" + path_name + "/发票图片")

        print(f"共计{len_of_df}, 已完成{one + 1}")