from docx.shared import Twips
from docx.text.run import Run

class Picture:
    def __init__(self, path: str, caption=None, width=Twips(5000), style=None):
        self.path = path
        self.caption = caption
        self.width = width
        self.style = style

    def addToDoc(self, run: Run):
        #img = Image.open(replace)
        #img_with_border = ImageOps.expand(img, border=1, fill='black')
        #imgByteArr = BytesIO()
        #img_with_border.save(imgByteArr, format='PNG')
        #replace = imgByteArr
        try:
            run.text = run.text.replace(search, "")
            run.add_picture(self.path, width=self.width)
        except Exception as e:
            print(f"Error adding picture, {self.path}")
            print(e)
