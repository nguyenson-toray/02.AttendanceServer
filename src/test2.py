import os

import cv2
from pyzbar.pyzbar import decode
from pdf2image import convert_from_path


def detect_qr_from_png(image_path):
    """
  Detects QR codes in a PNG image using OpenCV.

  Args:
      image_path (str): Path to the PNG image file.

  Returns:
      list: List of decoded QR code data (strings) or an empty list if none found.
  """
    qr_data = []

    try:
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale for better detection

        # Detect QR codes
        qr_codes = decode(gray)

        # Extract data from QR codes
        for qr_code in qr_codes:
            qr_data.append(qr_code.data.decode('utf-8'))

    except Exception as e:
        print(f"Error detecting QR code in {image_path}: {e}")
    return qr_data


def pdf_to_image(path: str) -> str:
    pages = convert_from_path(path, dpi=300, output_file='qr.png', paths_only=True, output_folder=r"D:\OT")
    # images.save(f'out.png', 'PNG')
    # for count, page in enumerate(images):
    #   page.save(f'out{count}.jpg', 'JPEG')
    for page in pages:
        detect_qr_from_png(page)


def detect_qr(folder_path: str):
    list_qr = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.pdf'):
            try:
                temp_pages = convert_from_path(file_path, dpi=200, output_file='qr.png', paths_only=True,
                                               output_folder=folder_path)
                for temp_page in temp_pages:
                    try:
                        page = cv2.imread(temp_page)
                        gray = cv2.cvtColor(page, cv2.COLOR_BGR2GRAY)  # Convert to grayscale for better detection
                        # Detect QR codes
                        qr_codes = decode(gray)
                        # Extract data from QR codes
                        for qr_code in qr_codes:
                            list_qr.append(qr_code.data.decode('utf-8'))
                    except Exception as e:
                        print(f"Error detecting QR code in {temp_page}: {e}")
            except Exception as e:
                print(f"Error Error detecting QR code in {filename}: {e}")
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.ppm'):
            os.remove(file_path)
    print(list_qr)
    return list_qr

detect_qr(r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\01.HR\20.OT request")