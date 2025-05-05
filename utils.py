import hashlib
import os
import shutil
import re
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
import logging
import sqlite3

try:
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    DEFAULT_FONT = 'STSong-Light'
except Exception as e:
    logging.error(f"字体加载失败：{str(e)}")
    DEFAULT_FONT = 'Helvetica'

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def check_password(input_password):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            c.execute("SELECT password_hash FROM users WHERE id=1")
            stored_hash = c.fetchone()
            if stored_hash:
                return hash_password(input_password) == stored_hash[0]
        return False
    except Exception as e:
        logging.error(f"密码验证失败：{str(e)}")
        return False

def save_password(new_password):
    try:
        with sqlite3.connect('hr_data.db') as conn:
            c = conn.cursor()
            c.execute("UPDATE users SET password_hash=? WHERE id=1", (hash_password(new_password),))
            if conn.total_changes == 0:
                c.execute("INSERT INTO users (id, password_hash) VALUES (1, ?)", (hash_password(new_password),))
            conn.commit()
        logging.info("密码已更新")
    except Exception as e:
        logging.error(f"保存密码失败：{str(e)}")
        raise

def validate_password(password):
    if len(password) < 8:
        return False, "密码长度需至少8位"
    if not re.search(r"[A-Za-z]", password):
        return False, "密码需包含字母"
    if not re.search(r"\d", password):
        return False, "密码需包含数字"
    if not re.search(r"[!@#$%^&*(),.?\":{}|<>]", password):
        return False, "密码需包含特殊字符"
    return True, "密码强度：强"

def upload_photo(file_path):
    if file_path:
        if not os.path.exists("photos"):
            os.makedirs("photos")
        new_path = f"photos/{os.path.basename(file_path)}"
        shutil.copy(file_path, new_path)
        logging.info(f"照片上传完成：{new_path}")
        return new_path
    return None

def delete_photo(photo_path_var, photo_label):
    photo_path_var.set("")
    photo_label.config(image='')
    photo_label.image = None
    tk.Label(photo_label, text="无照片", font=("Arial", 9)).pack(expand=True)

def backup_data(backup_path):
    if backup_path:
        shutil.copy('hr_data.db', backup_path)
        logging.info("数据备份完成")
        return "数据已备份！", None
    return None, "未选择备份路径"

def export_person_data(person, from_talent=False, file_path=None):
    if not file_path:
        logging.error("导出PDF失败：未提供文件路径")
        return None, "导出失败：未选择文件路径"
    try:
        c = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4
        margin = 20 * mm
        content_width = width - 2 * margin
        x = margin
        y = height - margin
        page_number = 1

        def new_page():
            nonlocal y, page_number
            c.setFont(DEFAULT_FONT, 9)
            c.drawRightString(width - margin, margin, f"第 {page_number} 页")
            c.showPage()
            page_number += 1
            y = height - margin

        # 仅首页绘制标题
        c.setFont(DEFAULT_FONT, 16)
        title_text = "个人简历" if DEFAULT_FONT != 'Helvetica' else "Personal Resume"
        c.drawCentredString(width / 2, y, title_text)
        y -= 15 * mm

        c.setFont(DEFAULT_FONT, 13)
        c.drawString(x, y, "基本信息")
        y -= 6 * mm
        c.setLineWidth(1)
        c.line(x, y, x + content_width, y)
        y -= 10 * mm

        photo_x = width - margin - 35 * mm
        photo_y = y - 35 * mm
        c.rect(photo_x, photo_y, 25 * mm, 35 * mm)
        has_photo = person[-1] and os.path.exists(person[-1])
        if has_photo:
            c.drawImage(person[-1], photo_x, photo_y, width=25 * mm, height=35 * mm, preserveAspectRatio=True)
        else:
            c.setFont(DEFAULT_FONT, 10)
            c.drawCentredString(photo_x + 12.5 * mm, photo_y + 17.5 * mm, "无照片")

        basic_fields = ['真实姓名', '性别', '年龄', '身份证号', '手机号', '分会职务', '在职状态']
        c.setFont(DEFAULT_FONT, 10)
        label_width = 25 * mm
        field_indices = [1, 2, 3, 4, 5, 13, 14]
        for i, field in enumerate(basic_fields):
            label = f"{field}："
            value = str(person[field_indices[i]]) if person[field_indices[i]] else "无"
            c.drawRightString(x + label_width, y, label)
            c.drawString(x + label_width + 2 * mm, y, value)
            y -= 8 * mm

        y -= 12.5 * mm
        c.setFont(DEFAULT_FONT, 13)
        c.drawString(x, y, "详细信息")
        y -= 6 * mm
        c.setLineWidth(1)
        c.line(x, y, x + content_width, y)
        y -= 10 * mm
        c.setFont(DEFAULT_FONT, 9)
        detail_fields = ['省份', '城市', '昵称', '学历', '政治面貌', '个人职业', '加入组织时间', '跟捐天数', '家庭住址']
        field_indices = [6, 7, 9, 10, 11, 12, 15, 16, 17]
        for i, field in enumerate(detail_fields):
            label = f"{field}："
            value = str(person[field_indices[i]]) if person[field_indices[i]] else "无"
            c.drawRightString(x + label_width, y, label)
            c.drawString(x + label_width + 2 * mm, y, value)
            y -= 8 * mm

        y -= 12.5 * mm
        c.setFont(DEFAULT_FONT, 13)
        c.drawString(x, y, "个人简历")
        y -= 6 * mm
        c.setLineWidth(1)
        c.line(x, y, x + content_width, y)
        y -= 10 * mm

        bio_text = person[-2] if person[-2] else "无"
        style = ParagraphStyle(name='Normal', fontName=DEFAULT_FONT, fontSize=9, leading=14)
        paragraphs = bio_text.split('\n')
        page_bottom = margin + 10 * mm
        for para_text in paragraphs:
            para = Paragraph(para_text, style)
            para_width, para_height = para.wrap(content_width, height)
            if y - para_height < page_bottom:
                new_page()
            para.drawOn(c, x, y - para_height)
            y -= para_height + 4 * mm

        if from_talent:
            y -= 12.5 * mm
            c.setFont(DEFAULT_FONT, 13)
            c.drawString(x, y, "加入人才库理由")
            y -= 6 * mm
            c.setLineWidth(1)
            c.line(x, y, x + content_width, y)
            y -= 10 * mm

            with sqlite3.connect('hr_data.db') as conn:
                c_db = conn.cursor()
                c_db.execute("SELECT reason FROM talent_pool WHERE person_id=?", (person[0],))
                reason = c_db.fetchone()
            reason_text = reason[0] if reason and reason[0] else "无"
            paragraphs = reason_text.split('\n')
            for para_text in paragraphs:
                para = Paragraph(para_text, style)
                para_width, para_height = para.wrap(content_width, height)
                if y - para_height < page_bottom:
                    new_page()
                para.drawOn(c, x, y - para_height)
                y -= para_height + 4 * mm

        c.setFont(DEFAULT_FONT, 9)
        c.drawRightString(width - margin, margin, f"第 {page_number} 页")
        c.showPage()
        c.save()
        logging.info(f"导出PDF完成：{file_path}")
        return "人员信息已导出为PDF！", None
    except Exception as e:
        logging.error(f"导出PDF失败：{str(e)}")
        return None, f"导出失败：{str(e)}"