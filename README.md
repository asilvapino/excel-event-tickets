# excel-event-tickets
excel event tickets generator in PDF with QR code and automatic email send.

#### PDF Tickets generator and mail send

Excel list of tickets id names and emails
QR Generator for tickets numbers
PDF with background image
PDF Attached email send with template

### install

pip3 install qrcode[pil] reportlab pandas

### configure at main func

excel_file > Path to your Excel file
background_image > Path to your background image (PDF Size according to the size of your image, Recommended: 1022x1022 px)
output_directory > Output directory for saving tickets
prefix_url > Prefix URL for the QR code

### usage

python3 app.py

### debug

email_errors.log

### author

Abner Silva Pino
Synergos SpA
