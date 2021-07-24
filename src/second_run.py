import sys
import os
import argparse
import pytesseract
from pyzbar.pyzbar import decode
from PIL import Image
from datetime import datetime
import jwt
import re
import pandas as pd
from collections import defaultdict
from jwt.exceptions import DecodeError
from collections import OrderedDict
from pyzbar.pyzbar import ZBarSymbol
import numpy as np
from file_path import *
from python_utils import get_logger
from functools import wraps
import time

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, Spacer, SimpleDocTemplate

logging = get_logger()
Image.MAX_IMAGE_PIXELS = None
DATETIME_FORMAT = "%m%d%y%H%M"


def timer(func):

    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        value = func(*args, **kwargs)
        end_time = time.perf_counter()
        run_time = end_time - start_time
        logging.info("total run time by func {} in {} sec".format(func.__name__, run_time))
        return value

    return wrapper


def check_dir_path():
    """
    function to check if required directories exist if not then exit.
    :return:
    """
    paths = [image_path, tesseract_path, public_key_path, poppler_path, output_path, undecoded_path, decoded_path, tif_path]
    for path in paths:
        if not os.path.exists(path):
            logging.info("{} doesn't exist. Please create the required directory and execute the script.".format(path))
            sys.exit(1)


def get_time_for_file():
    today = datetime.now().strftime(DATETIME_FORMAT)
    return today


def _get_public_key():
    """
    function to get e-invoice public key to decode the qr code.
    :return: returns public key
    """
    with open("{}".format(public_key_path), 'r') as pub:
        public_key = pub.read()
    return public_key


def _get_blank_data(result):
    """
    function to get blank data updated for non-qr and non-decoded invoices
    :param result:
    :return: updated the result dictionary with blank values
    """
    fields = ['Sellergstin', 'Buyergstin', 'Docno', 'Doctyp', 'Docdt', 'Totinvval', 'Itemcnt', 'Mainhsncode', 'Irn',
              'Irndt']
    for col in fields:
        result[col].append('')
    return result


def _extract_text_data(result, imagename, ponum_, qr=None):
    """
    function to extract PO/Vendor Name/Invoice No/Invoice Date/Total Amount from  input pdf->image files
    using pytesseract library
    :param result:
    :param imagename: input file name
    :param ponum_: PO number regex
    :param qr: qr/non qr file check
    :return: updates and return result dictionary
    """
    try:
        PONUM = ponum_
        pytesseract.pytesseract.tesseract_cmd = tesseract_path  # tessearact library to extract text from image
        text = str(
            (pytesseract.image_to_string(
                Image.open(r"{}\{}.tif".format(tif_path, imagename)))))
        # text_list = text.splitlines()
        text_list = [ll.rstrip() for ll in text.splitlines() if ll.strip()]
        # regular expression pattern for date and invoice amount
        date = re.compile(r'(\d{2}-\d{2}-\d{4})')
        date_str = re.compile(r'(\d{2}-\w{3}-\d{2})')
        total = re.compile(r'(\d{1,}\s+Nos)')
        po_available = False
        vendor_available = False
        for line in text_list:
            if "PO NO" in line.upper():
                PO = PONUM.findall(line)
                if PO:
                    PO = PO[0]
                    result['PO/NPO'].append(PO)
                    po_available = True
                    break
            elif (" PO " in line.upper() or "PO/" in line.upper()) and ":" in line:
                # PO = line.split(":")[1].split(" ")[2].split("/")[0]
                PO = PONUM.findall(line)
                if PO:
                    PO = PO[0]
                    result['PO/NPO'].append(PO)
                    po_available = True
                    break
        if not po_available:
            result['PO/NPO'].append('NON PO')

        # extract vendor name from invoice/image, will look for LTD/LIMITED string as there is no field as such to
        # for vendor name
        for line in text_list:
            if "LTD" in line.upper() and "TATA" not in line.upper():
                vend = re.compile(r'^(.*\sLtd)', re.IGNORECASE)
                VendorName = vend.findall(line)
                if VendorName:
                    VendorName = VendorName[0]
                    vendorsemi = re.findall(r'"([^"]*)', VendorName)
                    if vendorsemi:
                        result['Vendor Name'].append(vendorsemi[0])
                        vendor_available = True
                        break
                    elif "STARTRER LOGISTICS" in VendorName.upper() or "STARTREK LOGISTICS" in VendorName.upper():
                        result['Vendor Name'].append("SPOTON LOGisTics PvT. Ltd")
                        vendor_available = True
                        break
                    if ":" in VendorName:
                        VendorName = VendorName.split(":")[1]
                        result['Vendor Name'].append(VendorName)
                        vendor_available = True
                        break
                    result['Vendor Name'].append(VendorName)
                    vendor_available = True
                    break

            elif "LIMITED" in line.upper() and "TATA" not in line.upper():
                vend = re.compile(r'^(.*\sLimited)', re.IGNORECASE)
                VendorName = vend.findall(line)
                if VendorName:
                    VendorName = VendorName[0]
                    vendorsemi = re.findall(r'"([^"]*)', VendorName)
                    if vendorsemi:
                        result['Vendor Name'].append(vendorsemi[0])
                        vendor_available = True
                        break
                    elif "STARTRER LOGISTICS" in VendorName.upper():
                        result['Vendor Name'].append("SPOTON LOGisTics PvT. Ltd")
                        vendor_available = True
                        break
                    if ":" in VendorName:
                        VendorName = VendorName.split(":")[1]
                        result['Vendor Name'].append(VendorName)
                        vendor_available = True
                        break
                    result['Vendor Name'].append(VendorName)
                    vendor_available = True
                    break

        if not vendor_available:
            result['Vendor Name'].append('')
        if qr == "YES":
            return result

        invoice_date_avail = False
        invoice_no_avail = False
        invoice_amnt_avail = False

        for idx, line in enumerate(text_list):
            if "INVOICE NO / INVOICE DATE" in line.upper():
                invoice_no, invoice_date = line.split("Date")[1].split("|")
                result['Invoice No'].append(invoice_no)
                result['Invoice Date'].append(invoice_date)
                invoice_no_avail = True
                invoice_date_avail = True
                continue
            if "INVOICE NO" in line.upper():
                if ":" in line:
                    invoice_no = line.split(":")[1]
                else:
                    lin = text_list[idx + 1]
                    invoice_no = lin.strip() if lin else None
                result['Invoice No'].append(invoice_no)
                invoice_no_avail = True
                continue
            if "DATE" in line.upper():
                if "BILL DATE" in line.upper() or "DATED" in line.upper():
                    # invoice_date = line.split(":")[-1]
                    if line.upper() == "DATED":
                        lin = text_list[idx + 1].strip()
                        invoice_date = lin.strip() if lin else None
                        result['Invoice Date'].append(invoice_date)
                        invoice_date_avail = True
                else:
                    invoice_d = line.split(" ")[-1]
                    # print("invoice_d  val {}".format(invoice_d))
                    dat = date.findall(invoice_d)
                    dat_str = date_str.findall(invoice_d)
                    if dat:
                        result['Invoice Date'].append(dat)
                        invoice_date_avail = True
                    elif dat_str:
                        result['Invoice Date'].append(dat_str)
                        invoice_date_avail = True
                    continue
                continue
            if "TOTAL AMOUNT" in line.upper() or "Final Amount" in line.upper() or "GRAND TOTAL" in line.upper():
                invoice_amnt = line.split(" ")[-1]
                result['Invoice Amount'].append(invoice_amnt)
                invoice_amnt_avail = True

            elif total.findall(line):
                next_line = text_list[idx + 1].strip()
                if not next_line.startswith("="):
                    continue
                invoice_amnt = next_line.replace("=", '') if next_line else None
                #print("invoice amount 2 {}".format(invoice_amnt))
                result['Invoice Amount'].append(invoice_amnt)
                invoice_amnt_avail = True
            elif "ROUND OFF NET AMOUNT" in line.upper():
                invoice_amnt = line.split("amount")[-1]
                #print("invocie amnt 3 {}".format(invoice_amnt))
                result['Invoice Amount'].append(invoice_amnt)
                invoice_amnt_avail = True
        if not invoice_amnt_avail:
            result['Invoice Amount'].append('')
        if not invoice_date_avail:
            result['Invoice Date'].append('')
        if not invoice_no_avail:
            result['Invoice No'].append('')
        return result
    except Exception as e:
        logging.info("Exception raised while extracting text {}, {}, {}".format(e, imagename, sys.exc_info()[-1].tb_lineno))


def dataToParagraph(data):
    """
        function to read the QR code decoded data and build a story/list to generate pdf
        :param data: QR code data
        :return:
    """
    p = ""
    for key, val in data.items():
        if val:
            p += "<strong> {}: </strong>".format(key) + "    {}".format(val) + "<br/><br/>"
    return p


def _generate_pdf(path, pdf_data=None):
    """
      function to generate updated pdf where first page will be QR code decoded data
      :param path: pdf invoice/files path
      :param pdf_data: pdf/QR code decoded data
      :return: returns None.
    """
    try:
        story = []
        # define the style for our paragraph text
        styles = getSampleStyleSheet()
        styleN = styles['Normal']
        # keep on adding the first page QR code related data into story
        story.append(Paragraph("<strong>Results of E-invoice Decoder</strong>", styleN))
        story.append(Spacer(1, .25 * inch))

        # text = {"QR Code  Present (Y/N)": "No", "QR Code – Decode": "No QR Code",
        #         "Barcode": "1100454419", "Vendor Name": "Inflow Technologies Pvt. Ltd"}
        story.append(Paragraph(dataToParagraph(pdf_data), styleN))

        doc = SimpleDocTemplate(r"{}\temp_{}.pdf".format(temp_path, pdf_data['Barcode'])
                                , pagesize=letter, topMargin=0)
        doc.build(story)       # build the pdf doc based on the story list
        f_name = open(r"{}\temp_{}.pdf".format(temp_path, pdf_data['Barcode']), "rb")
        new_pdf = PdfFileReader(f_name, strict=False)

        if os.path.isfile(r"{}\{}.pdf".format(path, pdf_data['Barcode'])):
            e_name = open(r"{}\{}.pdf".format(path, pdf_data['Barcode']), "rb")
        else:
            im = Image.open(r"{}\{}_org.tif".format(tif_path, pdf_data['Barcode']))
            if im.mode == "RGBA":
               im = im.convert("RGB")
            im.save(r"{}\{}.pdf".format(tif_path, pdf_data['Barcode']), save_all=True)
            e_name = open(r"{}\{}.pdf".format(tif_path, pdf_data['Barcode']), "rb")

        # rgba = Image.open(r"{}\{}_org.tif".format(tif_path, pdf_data['Barcode']))
        # rgb = Image.new('RGB', rgba.size, (255, 255, 255))  # white background
        # rgb.paste(rgba, mask=rgba.split()[3])  # paste using alpha channel as mask
        # # rgb.save(PDF_FILE, 'PDF', resoultion=100.0)
        # rgb.save(r"{}\{}.pdf".format(tif_path, pdf_data['Barcode']), save_all=True)
        existing_pdf = PdfFileReader(e_name, strict=False)
        output = PdfFileWriter()
        output.addPage(new_pdf.getPage(0))
        pagecount = existing_pdf.getNumPages()
        for ind in range(pagecount):         # keep on adding/appending existing pdf pages to output object
            output.addPage(existing_pdf.getPage(ind))
        if pdf_data['QR Code – Decode'].upper() == "SUCCESS":  # updated pdf's will be saved under respective folders
            outputStream = open(r"{}\{}.pdf".format(decoded_path, pdf_data['Barcode']), "wb")
        elif pdf_data['QR Code – Decode'].upper() == "NO QR CODE":
            outputStream = open(r"{}\{}.pdf".format(undecoded_path, pdf_data['Barcode']), "wb")
        else:
            outputStream = open(r"{}\{}.pdf".format(qr_code_failure, pdf_data['Barcode']), "wb")
        output.write(outputStream)
        f_name.close()
        e_name.close()
        outputStream.close()
        os.remove(r"{}\temp_{}.pdf".format(temp_path, pdf_data['Barcode']))

    except Exception as e:
        logging.info("Exception raised while generating updated invoice pdfs: {}".format(e))


def remove_old_files(file_path):
    """
        function to delete old files from all the respective folders
        :param file_path: list of folder path's
        :return:
    """
    if os.path.isdir("{}".format(file_path)):
        files = os.listdir("{}".format(file_path))
        for file in files:
            os.remove(r"{}\{}".format(file_path, file))


def main():
    check_dir_path()       # check if all the folders are available or not
    remove_old_files(temp_path)

    public_key = _get_public_key()  # get public key to decode and validate QR code
    data = None
    PONUM = re.compile(r'\b[0-9]+\b')
    try:
        if os.path.isdir(tif_path):
            files = os.listdir(tif_path)
            result = defaultdict(list)
            if not files:
                logging.info("No tiff files under path {} please check. exiting..".format(tif_path))
                sys.exit(1)
            for file in files:
                logging.info("Extracting file {}".format(file))
                tiffile, f_type = file.split(".")
                if not f_type.upper() == "TIF":
                    continue
                try:
                    # pdfFileObj = open(r"{}\{}".format(path, tiffile), 'rb')
                    # pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)
                    # ret, barcodes = cv2.imreadmulti(r"{}\{}.tif".format(image_path, imagename),
                    #                                 flags=cv2.IMREAD_GRAYSCALE)
                    img = Image.open(r"{}\{}.tif".format(tif_path, tiffile))
                    IsQrCode = False
                    CorrectQr = False
                    n = 0
                    for ind in range(img.n_frames):  # read page by page if multiple pages
                        img.seek(ind)
                        img.save(r"{}\{}_{}.tif".format(tif_path, tiffile, ind))
                        QrCode = decode(Image.open(r"{}\{}_{}.tif".format(tif_path, tiffile, ind)),
                                        symbols=[ZBarSymbol.QRCODE])
                        if QrCode:
                            # print(QrCode)
                            # print("length of qr code {}".format(len(QrCode)))
                            IsQrCode = True
                            if len(QrCode) == 1:   # logic for one QR code or multiple Qr codes in a single page
                                try:
                                    jwt.decode(QrCode[0].data + "==".encode(), public_key, algorithms='RS256')
                                except DecodeError as _:
                                    n = ind
                                    continue
                                else:        # if valid QR code then rename the page to filename_org.tif so that later code can read exact file
                                    img.close()
                                    os.rename(r"{}\{}.tif".format(tif_path, tiffile),
                                              r"{}\{}_org.tif".format(tif_path, tiffile))
                                    os.rename(r"{}\{}_{}.tif".format(tif_path, tiffile, ind),
                                              r"{}\{}.tif".format(tif_path, tiffile))
                                    CorrectQr = True
                                    break
                            else:
                                for ix in range(len(QrCode)):
                                    try:
                                        jwt.decode(QrCode[ix].data + "==".encode(), public_key, algorithms='RS256')
                                    except DecodeError as _:
                                        n = ind
                                        continue
                                    else:
                                        img.close()
                                        os.rename(r"{}\{}.tif".format(tif_path, tiffile),
                                                  r"{}\{}_org.tif".format(tif_path, tiffile))

                                        os.rename(r"{}\{}_{}.tif".format(tif_path, tiffile, ind),
                                                  r"{}\{}.tif".format(tif_path, tiffile))
                                        CorrectQr = True
                                        break
                    # if no QR code then rename(backup) original tif file to filename_org.tif and first page of tif
                    # file(filename_0.tif) to filename.tif
                    if not IsQrCode and os.path.exists(r"{}\{}_0.tif".format(tif_path, tiffile)):
                        img.close()
                        os.rename(r"{}\{}.tif".format(tif_path, tiffile),
                                  r"{}\{}_org.tif".format(tif_path, tiffile))
                        os.rename(r"{}\{}_0.tif".format(tif_path, tiffile),
                                  r"{}\{}.tif".format(tif_path, tiffile))
                    # if QR code available but a invalid QR code then rename original tif to filename_org.tif
                    # and QR code page filename_{pagenumber}.tif to filename.tif (tif is for getting exact
                    # file to extract text data and generated updated pdf)
                    if IsQrCode and not CorrectQr:
                        img.close()
                        os.rename(r"{}\{}.tif".format(tif_path, tiffile),
                                  r"{}\{}_org.tif".format(tif_path, tiffile))
                        os.rename(r"{}\{}_{}.tif".format(tif_path, tiffile, n),
                                  r"{}\{}.tif".format(tif_path, tiffile))

                except Exception as e:
                    logging.info("Exception raised {} for file {}".format(e, tiffile))

                if os.path.exists(r"{}\{}.tif".format(tif_path, tiffile)):
                    QrCodeData = decode(Image.open(r"{}\{}.tif".format(tif_path, tiffile)), symbols=[ZBarSymbol.QRCODE])
                    #print(QrCodeData)
                    result['Received Date'].append('')
                    result['SOURCE (HARD/E-MAIL)'].append('')
                    result['Currency'].append('INR')
                    result['Barcode'].append("{}".format(tiffile))
                    #print("result {}".format(result))
                    if QrCodeData:
                        result['QR Code Present (Y/N)'].append("Yes")
                        logging.info("qr code available in {}".format(tiffile))
                    else:
                        logging.info("qr code not available in {}".format(tiffile))
                        # shutil.copy2(r"{}\{}".format(path, tiffile),
                        #              r"{}".format(undecoded_path))
                        result['QR Code Present (Y/N)'].append("No")
                        result['QR Code – Decode'].append("No QR Code")
                        pdf_data = {'QR Code Present (Y/N)': 'No', 'Barcode': tiffile}
                        result = _get_blank_data(result)
                        temp_result = _extract_text_data(result, tiffile, PONUM, "NO")
                        if temp_result:
                            result = temp_result
                        continue
                    temp_result = _extract_text_data(result, tiffile, PONUM, "YES")
                    if temp_result:
                        result = temp_result

                    for i in QrCodeData:
                        data = i.data + "==".encode()
                    try:
                        if len(QrCodeData) == 1:        # decode the QR code
                            decoded = jwt.decode(data, public_key, algorithms='RS256')
                        else:
                            for ix in range(len(QrCodeData)):  # decode multiple QR codes and skip the QR code if it's invalid
                                try:
                                    data = QrCodeData[ix].data + "==".encode()
                                    decoded = jwt.decode(data, public_key, algorithms='RS256')
                                except DecodeError as _:
                                    continue
                        # print(decoded)
                    except DecodeError as e:
                        result['QR Code – Decode'].append("FAILURE")
                        result = _get_blank_data(result)
                        # result = _extract_text_data(result, imagename, PONUM)
                        logging.info("DecodeError Exception {} for file {}".format(e, tiffile))
                        continue

                    result['QR Code – Decode'].append("SUCCESS")
                    payload_data = eval(decoded['data'])
                    for k, v in payload_data.items():    # read all the QR code decoded data into result dictionary
                        # print("{} : {}".format(k, v))
                        result[k.capitalize()].append(v)
            final_qr_data = OrderedDict()      # get all the required fields for output excel file
            fields = ['QR Code Present (Y/N)', 'QR Code – Decode', 'Received Date', 'Barcode',
                      'PO/NPO', 'Vendor Name', 'Invoice Date', 'Invoice No', 'Inv Gross Amt', 'Currency',
                      'SOURCE (HARD/E-MAIL)',
                      'Sellergstin', 'Buyergstin', 'Doctyp', 'Itemcnt', 'Mainhsncode',
                      'Irn', 'Irndt']
            for col in fields:            # rename few fields/column names from QR code fields to required excel columns
                if col == 'Invoice Date':
                    final_qr_data[col.capitalize()] = result.get('Docdt', None)
                    continue
                elif col == 'Invoice No':
                    final_qr_data[col] = result.get('Docno', None)
                    continue
                elif col == 'Inv Gross Amt':
                    final_qr_data[col] = result.get('Totinvval', None)
                    continue
                final_qr_data[col] = result.get(col, None)

            for ind in range(len(final_qr_data['QR Code – Decode'])):
                pdf_data = {}                          # generate updated pdf's
                for k, v in final_qr_data.items():
                    pdf_data[k] = v[ind]
                _generate_pdf(input_path, pdf_data)
            date_time = get_time_for_file()
            remove_old_files(tif_path)
            if len(result) > 0:
                df = pd.DataFrame(final_qr_data)
                df.index = np.arange(1, len(df) + 1)
                df.index.name = 'SN'
                df.to_csv(r'{}\QRcode details_{}.csv'.format(output_path, date_time), encoding='utf-8-sig')
            else:
                logging.info("Returned empty data. Excel file not generated.")
    except Exception as e:
        logging.info("On Exit received exception: {}".format(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
