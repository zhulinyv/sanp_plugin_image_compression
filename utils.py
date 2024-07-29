import os
import shutil
from pathlib import Path

import cv2
import openpyxl
from loguru import logger
from openpyxl.drawing.image import Image as OPXIMG
from openpyxl.styles import Alignment
from PIL import Image as PILIMG

from utils.imgtools import return_pnginfo, revert_img_info
from utils.utils import file_namel2pathl, file_path2list


def image_compression_(format_, image):
    cv2_image = cv2.imread(image)
    if format_ == "jpg":
        compression_params = [cv2.IMWRITE_JPEG_QUALITY, 90]
    elif format_ == "png":
        compression_params = [cv2.IMWRITE_PNG_COMPRESSION, 9]
    cv2.imwrite(f"./output/temp.{format_}", cv2_image, compression_params)
    if format_ == "png":
        with PILIMG.open(image) as pil_img:
            pnginfo = pil_img.info
            revert_img_info(None, f"./output/temp.{format_}", pnginfo)
    elif format_ == "jpg":
        pass
    os.remove(image)
    shutil.move(f"./output/temp.{format_}", str(image)[:-4] + f".{format_}")


def image_compression(format_, image_path):
    image_list = file_namel2pathl(file_path2list(image_path), image_path)
    for image in image_list:
        logger.info(f"正在压缩 {image} ...")
        image_compression_(format_, image)
    logger.success("压缩完成!")
    return "压缩完成!"


def image_organization(format_, image_path, switch):
    logger.info("正在整理图片...")
    image_list = file_namel2pathl(file_path2list(image_path), image_path)
    iamge_number = len(image_list)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.row_dimensions[1].height = 50
    for row in range(2, iamge_number + 2):
        ws.row_dimensions[row].height = 300
    for col in ["A", "B", "C"]:
        ws.column_dimensions[col].width = 40
    for col in range(4, 12):
        col_letter = openpyxl.utils.get_column_letter(col)
        ws.column_dimensions[col_letter].width = 20
    ws.append(
        [
            "图片",
            "正面提示词",
            "负面提示词",
            "分辨率",
            "采样步数",
            "提示词相关性",
            "噪声计划表",
            "采样器",
            "sm",
            "sm_dyn",
            "随机种子",
        ]
    )
    number = 2
    for image in image_list:
        with PILIMG.open(image) as pilimg:
            (
                positive_input,
                negative_input,
                resolution,
                steps,
                scale,
                noise_schedule,
                sampler,
                sm,
                sm_dyn,
                seed,
            ) = return_pnginfo(pilimg)
            w, h = pilimg.size
        if switch:
            image_compression_(format_, image)
            image = OPXIMG(str(image)[:-4] + f".{format_}")
        else:
            image = OPXIMG(image)
        image.width, image.height = 260, int(260 / w * h)
        ws.add_image(image, f"A{number}")
        ws[f"B{number}"] = positive_input
        ws[f"C{number}"] = negative_input
        ws[f"D{number}"] = resolution
        ws[f"E{number}"] = steps
        ws[f"F{number}"] = scale
        ws[f"G{number}"] = noise_schedule
        ws[f"H{number}"] = sampler
        ws[f"I{number}"] = sm
        ws[f"J{number}"] = sm_dyn
        ws[f"K{number}"] = seed
        number += 1
    alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = alignment
    wb.save(Path(image_path) / "organization.xlsx")
    logger.success("整理完成!")
    return "整理完成!"
