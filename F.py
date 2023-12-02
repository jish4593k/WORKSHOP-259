from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.comments import Comment

def convert_image_to_excel(image_path, output_excel_path, comment_text="Image Comment"):
    # Open the image
    img = Image.open(image_path)

    # Create a new workbook and select the active sheet
    wb = Workbook()
    ws = wb.active

    
    img.thumbnail(max_size)


    temp_image_path = "temp_image.png"
    img.save(temp_image_path)

    
    img_excel = ExcelImage(temp_image_path)
    ws.add_image(img_excel, 'A1')


    comment = Comment(comment_text, "Author")
    ws['A1'].comment = comment

    wb.save(output_excel_path)

    
    os.remove(temp_image_path)

if __name__ == "__main__":
    image_path = "path/to/your/image.jpg"
    output_excel_path = "output_image.xlsx"
    convert_image_to_excel(image_path, output_excel_path)
