from flask import Flask, render_template, request,send_file
from PIL import Image, ImageDraw, ImageFont
import xlrd
import io

app = Flask(__name__)

# Define the font to use for text
font = ImageFont.truetype('arial.ttf', 22)

# Define the position of each text element
positions = [
    (260, 400),
    (1760, 400),
    (940, 555),
    (1400, 555),
    (1800, 555),
    (500, 505)
]

# Load the blank certificate image
certificate_blank = Image.open('Print_Marksheet.tiff')

# Open the Excel sheet
workbook = xlrd.open_workbook('Registration_Details.xls')
sheet = workbook.sheet_by_index(0)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    # Get the data from the form
    name = request.form['name']
    rollno = request.form['rollno']

    # Find the row in the Excel sheet
    row = None
    for i in range(1, sheet.nrows):
        if sheet.cell_value(i, 1) == name and sheet.cell_value(i, 3) == rollno:
            row = i
            break

    # If the row is found, generate the marksheet
    if row is not None:
        # Get the data for this row
        name = sheet.cell_value(row, 1)
        date = sheet.cell_value(row, 2)
        rollno = sheet.cell_value(row, 3)
        enroll = sheet.cell_value(row, 4)
        batch = sheet.cell_value(row, 5)
        fname = sheet.cell_value(row, 6)

        # Combine the text elements into a single string
        text = f'{name}\n{date}\n{rollno}\n{enroll}\n{batch}\n{fname}'

        # Create a new image with the text added
        certificate = certificate_blank.copy()
        draw = ImageDraw.Draw(certificate)
        for position, line in zip(positions, text.split('\n')):
            draw.text(position, line, font=font, fill=(0, 0, 0))
 # Save the image as a bytes buffer
         # Save the image as bytes
            img_bytes = io.BytesIO()
            certificate.save(img_bytes, format='TIFF')
            img_bytes.seek(0)

            return send_file(io.BytesIO(img_bytes.getvalue()), mimetype='image/tif', as_attachment=True, download_name='marksheet.tif')


    else:
        return 'Data not found'

if __name__ == '__main__':
    app.run(debug=True)
