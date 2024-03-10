import docx
from PIL import Image
import pytesseract
import io
import os

# 配置Tesseract的路径
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def save_image_if_text_found(docx_path, output_folder, query, lang='chi_sim'):
    """从Word文档中提取图片，如果图片包含指定文字则保存图片"""
    doc = docx.Document(docx_path)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for i, rel in enumerate(doc.part.rels.values(), start=1):
        if "image" in rel.target_ref:
            img_blob = rel.target_part.blob
            image = Image.open(io.BytesIO(img_blob))
            text = pytesseract.image_to_string(image, lang=lang)
            if query in text:
                image_path = os.path.join(output_folder, f"image_{i}.png")
                image.save(image_path)
                print(f"Saved '{image_path}' as it contains the query text.")

# 使用示例
docx_path = 'test.docx'  # 替换为你的Word文档路径
output_folder = './extracted_images'  # 设置图片保存的文件夹
query = '零点'  # 设置你想要在图片中查找的文字
save_image_if_text_found(docx_path, output_folder, query, lang='chi_sim')
