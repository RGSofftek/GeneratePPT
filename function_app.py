"""
Módulo de Azure Function que genera una presentación PowerPoint integrando visualizaciones de datos.
Descarga archivos Excel y una plantilla PowerPoint desde Azure File Share, genera gráficos, los inserta en diapositivas dinámicas,
sube la presentación generada y devuelve una URL pública.
"""

import azure.functions as func
import logging
import os
import json
import time
from datetime import datetime
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from io import BytesIO
from azure.storage.fileshare import ShareServiceClient
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from PIL import Image

# Configuración de variables de entorno
STORAGE_ACCOUNT_NAME = os.environ.get("STORAGE_ACCOUNT_NAME")
SAS_TOKEN = os.environ.get("SAS_TOKEN")
AUTH_TOKEN_EXPECTED = os.environ.get("AUTH_TOKEN")
FILE_SHARE_NAME = os.environ.get("FILE_SHARE_NAME")
DIRECTORY_NAME = os.environ.get("DIRECTORY_NAME")
TEMPLATES_DIRECTORY = f"{DIRECTORY_NAME}/templates"
INPUTS_DIRECTORY = f"{DIRECTORY_NAME}/inputs"
OUTPUTS_DIRECTORY = f"{DIRECTORY_NAME}/outputs"

service_url = f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net"
SERVICE_CLIENT = ShareServiceClient(account_url=service_url, credential=SAS_TOKEN)
SHARE_CLIENT = SERVICE_CLIENT.get_share_client(FILE_SHARE_NAME)

def download_file_from_share(temp_dir: str, filename: str, folder: str) -> str:
    """Descarga un archivo desde Azure File Share y lo guarda localmente."""
    directory_client = SHARE_CLIENT.get_directory_client(folder)
    file_client = directory_client.get_file_client(filename)
    stream = file_client.download_file()
    local_file_path = os.path.join(temp_dir, filename)
    with open(local_file_path, 'wb') as f:
        f.write(stream.readall())
    logging.info(f"Descargado '{filename}' desde '{folder}' a '{local_file_path}'")
    return local_file_path

def upload_file_to_share(local_file_path: str, filename: str, folder: str) -> None:
    """Sube un archivo local a Azure File Share."""
    directory_client = SHARE_CLIENT.get_directory_client(folder)
    file_client = directory_client.get_file_client(filename)
    try:
        file_client.delete_file()
    except Exception:
        pass
    with open(local_file_path, 'rb') as data:
        file_client.upload_file(data)
    logging.info(f"Subido archivo '{filename}' a '{folder}'")

def generate_file_url(filename: str, folder: str) -> str:
    """Genera una URL pública para un archivo en Azure File Share."""
    return f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net/{FILE_SHARE_NAME}/{folder}/{filename}?{SAS_TOKEN}"

def retry_call(func_call, attempts: int = 2, delay: int = 1):
    """Ejecuta una función con lógica de reintentos."""
    last_exception = None
    for attempt in range(attempts):
        try:
            return func_call()
        except Exception as e:
            last_exception = e
            logging.error(f"Intento {attempt + 1} falló: {e}")
            time.sleep(delay)
    raise last_exception

def create_dynamic_slide(prs: Presentation, images: list, title_text: str = None):
    """
    Crea una diapositiva dinámica sin usar placeholders existentes:
    - Agrega título, imágenes y análisis usando textboxes e imágenes personalizadas.
    """
    blank_layout = prs.slide_layouts[6]  # Layout "Blank"
    slide = prs.slides.add_slide(blank_layout)

    # Definir márgenes y alturas fijas
    margin = Inches(0.5)
    title_height = Inches(1) if title_text else 0
    analysis_height = Inches(1)
    available_height = prs.slide_height - 2 * margin - title_height - analysis_height
    available_width = prs.slide_width - 2 * margin

    # Insertar título como textbox
    if title_text:
        txBox = slide.shapes.add_textbox(
            left=margin,
            top=margin,
            width=available_width,
            height=title_height
        )
        tf = txBox.text_frame
        tf.text = title_text
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(24)
                run.font.bold = True

    # Calcular espacio para imágenes
    images_top = margin + title_height
    images_left = margin

    # Determinar rejilla según el número de imágenes
    num_images = len(images)
    if num_images == 1:
        rows, cols = 1, 1
    elif num_images <= 2:
        rows, cols = 1, num_images
    elif num_images <= 4:
        rows, cols = 2, 2
    else:
        cols = int(num_images ** 0.5)
        rows = cols if cols * cols >= num_images else cols + 1

    cell_width = available_width / cols
    cell_height = available_height / rows

    # Insertar imágenes con escalado
    for i, img_stream in enumerate(images):
        col = i % cols
        row = i // cols
        cell_left = images_left + col * cell_width
        cell_top = images_top + row * cell_height

        img_stream.seek(0)
        img = Image.open(img_stream)
        orig_w, orig_h = img.size
        img_stream.seek(0)

        if (cell_width / cell_height) > (orig_w / orig_h):
            picture_height = cell_height
            picture_width = cell_height * (orig_w / orig_h)
        else:
            picture_width = cell_width
            picture_height = cell_width * (orig_h / orig_w)

        picture_left = cell_left + (cell_width - picture_width) / 2
        picture_top = cell_top + (cell_height - picture_height) / 2

        slide.shapes.add_picture(img_stream, picture_left, picture_top, width=picture_width, height=picture_height)

    # Insertar análisis como textbox
    analysis_text = "Este es el análisis generado para la gráfica."
    analysisBox = slide.shapes.add_textbox(
        left=margin,
        top=prs.slide_height - margin - analysis_height,
        width=available_width,
        height=analysis_height
    )
    analysis_tf = analysisBox.text_frame
    analysis_tf.text = analysis_text
    for paragraph in analysis_tf.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(80, 80, 80)

    return slide

def generate_kpr_graph(tmd_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> BytesIO:
    df_users = pd.read_excel(users_file_path)
    df_tmd = pd.read_excel(tmd_file_path)
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    df_filtered = df_tmd[df_tmd["Matricula LT"].isin(users_list)]
    plt.figure(figsize=(10, 6))
    plt.bar(df_filtered["Lider Técnico"], df_filtered[q], color='#000474')
    plt.title(f'{q} Values by Selected Lider Técnico (KPR)', fontsize=16)
    plt.xlabel('Lider Técnico', fontsize=12)
    plt.ylabel(q, fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()
    return img_bytes

def generate_kr_graph(pases_file_path: str, revisiones_file_path: str) -> BytesIO:
    quarters = ['Q1 01', 'Q1 02', 'Q1 03', 'Q2 01', 'Q2 02', 'Q2 03']
    df1 = pd.read_excel(pases_file_path)
    df2 = pd.read_excel(revisiones_file_path)
    df1_grouped = df1[quarters].sum(axis=0)
    df2_grouped = df2[quarters].sum(axis=0)
    plt.figure(figsize=(10, 7))
    x = range(len(quarters))
    plt.bar([pos - 0.2 for pos in x], df1_grouped, color='#000474', width=0.4, label='Dataset 1')
    plt.bar([pos + 0.2 for pos in x], df2_grouped, color='#f46a34', width=0.4, label='Dataset 2')
    plt.title('Pases Exitosos vs Reversiones x Q (KR)', fontsize=16)
    plt.xlabel('Trimestre', fontsize=12)
    plt.ylabel('Valor Total', fontsize=12)
    plt.xticks(ticks=x, labels=quarters)
    plt.legend()
    plt.tight_layout()
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()
    return img_bytes

def generate_maturity_graphs(maturity_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> list:
    df_users = pd.read_excel(users_file_path)
    df_maturity = pd.read_excel(maturity_file_path)
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    df_filtered = df_maturity[df_maturity["Matricula LT"].isin(users_list)]
    practices = df_filtered['PRACTICA'].unique()
    maturity_images = []
    for practice in practices[:4]:
        df_practice = df_filtered[df_filtered['PRACTICA'] == practice]
        plt.figure(figsize=(10, 6))
        plt.bar(df_practice['Lider Técnico'], df_practice[q], color='#000474')
        plt.title(f'Niveles de Madurez para {practice}', fontsize=16)
        plt.xlabel('Lider Técnico', fontsize=12)
        plt.ylabel(f'Nivel de Madurez ({q})', fontsize=12)
        plt.xticks(rotation=45)
        plt.tight_layout()
        img_bytes = BytesIO()
        plt.savefig(img_bytes, format='png')
        img_bytes.seek(0)
        plt.close()
        maturity_images.append(img_bytes)
    return maturity_images

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.function_name(name="generate_presentation")
@app.route(route="generate_presentation", methods=["POST"])
def generate_presentation(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Procesando solicitud de generate_presentation.")
    auth_header = req.headers.get("Authorization")
    if not auth_header or auth_header != AUTH_TOKEN_EXPECTED:
        return func.HttpResponse("No autorizado", status_code=401)
    
    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse("JSON inválido en el cuerpo de la solicitud.", status_code=400)
    
    required_keys = ["q", "matricula_lider", "tmd_file", "users_file", "pases_file", "revisiones_file", "maturity_level_file"]
    missing_keys = [key for key in required_keys if key not in req_body]
    if missing_keys:
        return func.HttpResponse(f"Faltan claves requeridas: {missing_keys}", status_code=400)
    
    q = req_body["q"]
    matricula_lider = req_body["matricula_lider"]
    tmd_file = req_body["tmd_file"]
    users_file = req_body["users_file"]
    pases_file = req_body["pases_file"]
    revisiones_file = req_body["revisiones_file"]
    maturity_level_file = req_body["maturity_level_file"]
    
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            tmd_file_path = download_file_from_share(temp_dir, tmd_file, INPUTS_DIRECTORY)
            users_file_path = download_file_from_share(temp_dir, users_file, INPUTS_DIRECTORY)
            pases_file_path = download_file_from_share(temp_dir, pases_file, INPUTS_DIRECTORY)
            revisiones_file_path = download_file_from_share(temp_dir, revisiones_file, INPUTS_DIRECTORY)
            maturity_file_path = download_file_from_share(temp_dir, maturity_level_file, INPUTS_DIRECTORY)
            template_file_path = download_file_from_share(temp_dir, "base_template.pptx", TEMPLATES_DIRECTORY)
            
            kpr_img = retry_call(lambda: generate_kpr_graph(tmd_file_path, users_file_path, matricula_lider, q))
            kr_img = retry_call(lambda: generate_kr_graph(pases_file_path, revisiones_file_path))
            maturity_imgs = retry_call(lambda: generate_maturity_graphs(maturity_file_path, users_file_path, matricula_lider, q))
            
            prs = Presentation(template_file_path)
            create_dynamic_slide(prs, images=[kpr_img], title_text=f"KPR - {q}")
            create_dynamic_slide(prs, images=[kr_img], title_text="KR")
            create_dynamic_slide(prs, images=maturity_imgs, title_text="Niveles de Madurez")
            
            file_date = datetime.now().strftime("%m%d%y")
            output_filename = f"{file_date} Presentation.pptx"
            output_file_path = os.path.join(temp_dir, output_filename)
            prs.save(output_file_path)
            
            upload_file_to_share(output_file_path, output_filename, OUTPUTS_DIRECTORY)
            public_url = generate_file_url(output_filename, OUTPUTS_DIRECTORY)
            
            return func.HttpResponse(json.dumps({"public_url": public_url}), status_code=200, mimetype="application/json")
        
        except Exception as e:
            logging.error(f"Error en generate_presentation: {e}")
            return func.HttpResponse(f"Ocurrió un error: {str(e)}", status_code=500)