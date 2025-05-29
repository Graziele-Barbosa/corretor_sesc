import os
import re
import json
from pdf2image import convert_from_path
import cv2
import numpy as np
from pyzbar.pyzbar import decode
from unidecode import unidecode

poppler_path = r'E:\Programas\Poopler\poppler-24.07.0\Library\bin'

def sanitize_filename(name):
    name = unidecode(name)
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.strip()
    return name


def pdf_to_images(pdf_path, output_dir):
    pages = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
    for i, page in enumerate(pages):
        temp_output_file = f'{output_dir}/temp_page_{i + 1}.png'
        page.save(temp_output_file, 'PNG')
        qr_code_name = get_qr_code_name(temp_output_file)
        if qr_code_name:
            sanitized_name = sanitize_filename(qr_code_name)
            new_image_path = os.path.join(output_dir, f"{sanitized_name}.png")
            os.rename(temp_output_file, new_image_path)
            print(f"Imagem renomeada para {new_image_path}")
        else:
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            page_number = f"page_{i + 1}"
            new_image_path = os.path.join(output_dir, f"{pdf_name}_{page_number}.png")
            print(f"QR code não encontrado na página {i + 1}, imagem renomeada para {new_image_path}")
            os.rename(temp_output_file, new_image_path)
    print(f"PDF convertido em imagens e salvo em: {output_dir}")


def get_qr_code_name(image_path):
    image = cv2.imread(image_path)
    if image is None:
        return None
    decoded_objects = decode(image)
    for obj in decoded_objects:
        if obj.type == 'QRCODE':
            qr_data = obj.data.decode('utf-8')
            try:
                qr_json = json.loads(qr_data)
                return qr_json.get('nome', '')
            except json.JSONDecodeError:
                print(f"Erro ao tentar parsear o QR code: {qr_data}")
                return None
    return None


def order_points(pts):
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect


def find_ref_point_in_roi(image, roi, expected_corner):

    x_start, y_start, width, height = roi
    cropped_image = image[y_start:y_start + height, x_start:x_start + width]
    gray = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    ret, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.dilate(thresh, kernel, iterations=1)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    candidates = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < 100:
            continue
        epsilon = 0.02 * cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, epsilon, True)
        if len(approx) == 4:
            x, y, w, h = cv2.boundingRect(approx)
            aspect_ratio = float(w) / h
            if 0.8 < aspect_ratio < 1.2:
                candidate = [x + x_start, y + y_start]
                candidates.append(candidate)
    if candidates:
        candidates = np.array(candidates)
        expected = np.array(expected_corner)
        distances = np.linalg.norm(candidates - expected, axis=1)
        best_candidate = candidates[np.argmin(distances)]
        return best_candidate.tolist()
    else:
        return None

def align_image(image_path, output_folder):
    gabarito = cv2.imread(image_path)
    if gabarito is None:
        return
    height, width = gabarito.shape[:2]
    roi_size = 500
    rois = [
        ((0, 0, roi_size, roi_size), (0, 0)),
        ((width - roi_size, 0, roi_size, roi_size), (width, 0)),
        ((0, height - roi_size, roi_size, roi_size), (0, height)),
        ((width - roi_size, height - roi_size, roi_size, roi_size), (width, height))
    ]
    ref_points = []
    for roi, expected_corner in rois:
        point = find_ref_point_in_roi(gabarito, roi, expected_corner)
        if point:
            ref_points.append(point)
    if len(ref_points) == 4:
        ordered_points = order_points(np.array(ref_points, dtype="float32"))
        dst_points = np.array([
            [0, 0],
            [width - 1, 0],
            [width - 1, height - 1],
            [0, height - 1]
        ], dtype="float32")
        M = cv2.getPerspectiveTransform(ordered_points, dst_points)
        aligned_gabarito = cv2.warpPerspective(gabarito, M, (width, height))
        resized_image = cv2.resize(aligned_gabarito, (1654, 2339))
        output_path = os.path.join(output_folder, os.path.basename(image_path))
        cv2.imwrite(output_path, resized_image)
        print(f"Processed and resized {os.path.basename(image_path)}")
    else:
        print(f"Error: Not enough reference points found in {os.path.basename(image_path)}")


def process_pdfs_and_align_images(base_dir):
    print("Escolha uma pasta para processar:")
    folders = [folder for folder in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, folder))]
    for idx, folder in enumerate(folders, 1):
        print(f"{idx}. {folder}")
    chosen_folder_index = int(input("Digite o número da pasta: ")) - 1
    chosen_folder = folders[chosen_folder_index]
    folder_path = os.path.join(base_dir, chosen_folder)
    print(f"Processando a pasta: {chosen_folder}")
    for file in os.listdir(folder_path):
        if file.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, file)
            print(f"Convertendo PDF: {pdf_path}")
            pdf_to_images(pdf_path, folder_path)
            for image_file in os.listdir(folder_path):
                if image_file.endswith('.png'):
                    image_path = os.path.join(folder_path, image_file)
                    align_image(image_path, folder_path)


base_dir = '03_pdf_to_image'
process_pdfs_and_align_images(base_dir)
print("Processo completo.")
