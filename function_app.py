"""
This Azure Function module generates a PowerPoint presentation by integrating several data visualizations.
It downloads input Excel files and a PowerPoint template from an Azure File Share, generates graphs based on 
the Excel data, inserts them into dynamically created slides, uploads the final presentation back to the share,
and returns a public URL for the generated PPT.

Expected JSON Request Body:
{
  "q": "<quarter_or_metric>",
  "matricula_lider": "<leader_identifier>",
  "tmd_file": "<TMD Excel filename>",
  "users_file": "<Users Excel filename>",
  "pases_file": "<Pases Excel filename>",
  "revisiones_file": "<Revisiones Excel filename>",
  "maturity_level_file": "<Maturity Excel filename>"
}

Authentication:
An "Authorization" header with a token matching the AUTH_TOKEN environment variable is required.
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
from PIL import Image  # Para obtener dimensiones originales (opcional)

# Environment variables and storage structure definitions
STORAGE_ACCOUNT_NAME = os.environ.get("STORAGE_ACCOUNT_NAME")
SAS_TOKEN = os.environ.get("SAS_TOKEN")
AUTH_TOKEN_EXPECTED = os.environ.get("AUTH_TOKEN")  # Expected token for authentication

FILE_SHARE_NAME = os.environ.get("FILE_SHARE_NAME")  # Name of the Azure File Share
DIRECTORY_NAME = os.environ.get("DIRECTORY_NAME")      # Root directory for the function
TEMPLATES_DIRECTORY = f"{DIRECTORY_NAME}/templates"     # Folder for PowerPoint templates
INPUTS_DIRECTORY = f"{DIRECTORY_NAME}/inputs"           # Folder for Excel input files
OUTPUTS_DIRECTORY = f"{DIRECTORY_NAME}/outputs"         # Folder for generated presentations

# Initialize the Azure File Share client
service_url = f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net"
SERVICE_CLIENT = ShareServiceClient(account_url=service_url, credential=SAS_TOKEN)
SHARE_CLIENT = SERVICE_CLIENT.get_share_client(FILE_SHARE_NAME)

def download_file_from_share(temp_dir: str, filename: str, folder: str) -> str:
    """
    Downloads a file from the specified Azure File Share folder and saves it locally.
    """
    directory_client = SHARE_CLIENT.get_directory_client(folder)
    file_client = directory_client.get_file_client(filename)
    stream = file_client.download_file()
    file_content = stream.readall()
    local_file_path = os.path.join(temp_dir, filename)
    with open(local_file_path, 'wb') as f:
        f.write(file_content)
    logging.info(f"Downloaded '{filename}' from '{folder}' to '{local_file_path}'")
    return local_file_path

def upload_file_to_share(local_file_path: str, filename: str, folder: str) -> None:
    """
    Uploads a local file to the specified folder in the Azure File Share.
    """
    directory_client = SHARE_CLIENT.get_directory_client(folder)
    file_client = directory_client.get_file_client(filename)
    try:
        file_client.delete_file()
        logging.info(f"Deleted existing file '{filename}' in folder '{folder}'.")
    except Exception as e:
        logging.info(f"File '{filename}' did not exist or could not be deleted: {e}")
    
    with open(local_file_path, 'rb') as data:
        file_client.upload_file(data)
    logging.info(f"Uploaded file '{filename}' to folder '{folder}'")

def generate_file_url(filename: str, folder: str) -> str:
    """
    Constructs a public URL for a file in the Azure File Share using the SAS token.
    """
    return f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net/{FILE_SHARE_NAME}/{folder}/{filename}?{SAS_TOKEN}"

def retry_call(func_call, attempts: int = 2, delay: int = 1):
    """
    Executes a function with retry logic in case of transient errors.
    """
    last_exception = None
    for attempt in range(attempts):
        try:
            return func_call()
        except Exception as e:
            last_exception = e
            logging.error(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(delay)
    raise last_exception

def insert_picture_on_slide(slide, image: BytesIO, position: tuple) -> None:
    """
    Inserts an image into a PowerPoint slide at the specified position.
    Position is a tuple (left, top, width, height) using pptx.util units.
    """
    left, top, width, height = position
    slide.shapes.add_picture(image, left, top, width=width, height=height)

# --- New functions for dynamic layout ---

def calculate_dynamic_layout(slide_width, slide_height, num_images, include_title=True, include_analysis=True):
    """
    Calculates and returns:
      - title_area: (left, top, width, height) for the title (if applicable)
      - image_positions: list of positions (left, top, width, height) for each image
      - analysis_area: (left, top, width, height) for the analysis block
    Based on the total slide dimensions and the number of images.
    """
    margin = Inches(0.5)
    usable_width = slide_width - 2 * margin
    usable_height = slide_height - 2 * margin

    title_area = None
    analysis_area = None

    title_height = Inches(1) if include_title else 0
    analysis_height = Inches(1.5) if include_analysis else 0

    images_area_top = margin + title_height
    images_area_height = usable_height - title_height - analysis_height

    # Determine grid based on number of images
    if num_images == 1:
        rows, cols = 1, 1
    elif num_images <= 2:
        rows, cols = 1, num_images
    elif num_images <= 4:
        rows, cols = 2, 2
    else:
        cols = int(num_images ** 0.5)
        rows = cols if cols * cols >= num_images else cols + 1

    cell_width = usable_width / cols
    cell_height = images_area_height / rows

    image_positions = []
    for i in range(num_images):
        col = i % cols
        row = i // cols
        left = margin + col * cell_width
        top = images_area_top + row * cell_height
        image_positions.append((left, top, cell_width, cell_height))

    if include_title:
        title_area = (margin, margin, usable_width, title_height)
    if include_analysis:
        analysis_area = (margin, margin + title_height + images_area_height, usable_width, analysis_height)

    return title_area, image_positions, analysis_area

def get_image_dimensions(image_stream: BytesIO):
    """
    Uses Pillow to obtain the original dimensions of the image.
    Returns a tuple (width, height).
    """
    image = Image.open(image_stream)
    return image.size  # (width, height)

def get_analysis_text(image_info: dict) -> str:
    """
    Simulates a synchronous call to OpenAI to obtain an analysis based on the graph(s) information.
    'image_info' may contain data such as type, metrics, etc.
    """
    # Here you would integrate the actual call to OpenAI.
    return "Este es el análisis generado para la gráfica, resaltando tendencias y puntos clave."

def create_dynamic_slide(prs: Presentation, images: list, title_text: str = None):
    """
    Creates a new dynamic slide:
      - Inserts a title (if provided).
      - Dynamically distributes all images.
      - Inserts the analysis block obtained from OpenAI.
    """
    blank_layout = prs.slide_layouts[6]  # Using the "Blank" layout for flexibility
    slide = prs.slides.add_slide(blank_layout)
    
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Calculate dynamic layout areas
    title_area, image_positions, analysis_area = calculate_dynamic_layout(
        slide_width, slide_height, len(images),
        include_title=(title_text is not None),
        include_analysis=True
    )

    # Insert title if provided
    if title_text and title_area:
        txBox = slide.shapes.add_textbox(*title_area)
        tf = txBox.text_frame
        tf.text = title_text
        # Optional formatting
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(24)
                run.font.bold = True

    # Insert images at calculated positions
    for img_stream, pos in zip(images, image_positions):
        # Optional: adjust scaling using get_image_dimensions to preserve aspect ratio
        slide.shapes.add_picture(img_stream, *pos)

    # Get and insert analysis text synchronously
    analysis_text = get_analysis_text({"num_images": len(images)})
    if analysis_area:
        analysisBox = slide.shapes.add_textbox(*analysis_area)
        analysis_tf = analysisBox.text_frame
        analysis_tf.text = analysis_text
        # Optional formatting
        for paragraph in analysis_tf.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(80, 80, 80)
    
    return slide

# --- Existing graph generation functions (KPR, KR, Maturity) remain unchanged ---

def generate_kpr_graph(tmd_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> BytesIO:
    df_users = pd.read_excel(users_file_path)
    required_users_cols = {"Matricula Lider", "Matricula"}
    if not required_users_cols.issubset(set(df_users.columns)):
        raise ValueError(f"Users file is missing required columns: {required_users_cols - set(df_users.columns)}")
    
    df_tmd = pd.read_excel(tmd_file_path)
    required_tmd_cols = {"Matricula LT", "Lider Técnico", q}
    if not required_tmd_cols.issubset(set(df_tmd.columns)):
        raise ValueError(f"TMD file is missing required columns: {required_tmd_cols - set(df_tmd.columns)}")
    
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    if not users_list:
        raise ValueError("No users found matching the provided 'Matricula Lider' in the Users file.")
    
    df_filtered = df_tmd[df_tmd["Matricula LT"].isin(users_list)]
    if df_filtered.empty:
        raise ValueError("No matching 'Lider Técnico' data found in TMD file for KPR graph.")
    
    plt.figure(figsize=(10, 6))
    bars = plt.bar(df_filtered["Lider Técnico"], df_filtered[q], color='#000474')
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval + 0.5, f'{yval:.2f}', ha='center', va='bottom', fontsize=10, color='black', zorder=3)
    
    plt.title(f'{q} Values by Selected Lider Técnico (KPR)', fontsize=16)
    plt.xlabel('Lider Técnico', fontsize=12)
    plt.ylabel(q, fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
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
    
    missing_df1 = set(quarters) - set(df1.columns)
    missing_df2 = set(quarters) - set(df2.columns)
    if missing_df1:
        raise ValueError(f"Pases file is missing required columns: {missing_df1}")
    if missing_df2:
        raise ValueError(f"Revisiones file is missing required columns: {missing_df2}")
    
    df1_grouped = df1[quarters].sum(axis=0)
    df2_grouped = df2[quarters].sum(axis=0)

    bar_width = 0.4
    x = range(len(quarters))
    x1 = [pos - bar_width / 2 for pos in x]
    x2 = [pos + bar_width / 2 for pos in x]

    plt.figure(figsize=(10, 7))
    bars1 = plt.bar(x1, df1_grouped, color='#000474', width=bar_width, label='Dataset 1', edgecolor='black', zorder=2)
    bars2 = plt.bar(x2, df2_grouped, color='#f46a34', width=bar_width, label='Dataset 2', edgecolor='black', zorder=1)
    
    for bar in bars1:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval, round(yval, 2), va='bottom', ha='center', fontsize=10, color='black', zorder=3)
    for bar in bars2:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval, round(yval, 2), va='bottom', ha='center', fontsize=10, color='black', zorder=3)
    
    plt.title('Pases Exitosos vs Reversiones x Q (KR)', fontsize=16)
    plt.xlabel('Trimestre', fontsize=12)
    plt.ylabel('Valor Total', fontsize=12)
    plt.xticks(ticks=x, labels=quarters, rotation=0)
    plt.grid(linestyle='--', linewidth=0.5)
    plt.legend()
    plt.tight_layout()

    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()
    return img_bytes

def generate_maturity_graphs(maturity_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> list:
    df_users = pd.read_excel(users_file_path)
    required_users_cols = {"Matricula Lider", "Matricula"}
    if not required_users_cols.issubset(set(df_users.columns)):
        raise ValueError(f"Users file is missing required columns: {required_users_cols - set(df_users.columns)}")
    
    df_maturity = pd.read_excel(maturity_file_path)
    required_maturity_cols = {"Matricula LT", "PRACTICA", "Lider Técnico", q}
    if not required_maturity_cols.issubset(set(df_maturity.columns)):
        raise ValueError(f"Maturity file is missing required columns: {required_maturity_cols - set(df_maturity.columns)}")
    
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    if not users_list:
        raise ValueError("No users found matching the provided 'Matricula Lider' in the Users file.")
    
    df_filtered = df_maturity[df_maturity["Matricula LT"].isin(users_list)]
    if df_filtered.empty:
        raise ValueError("No matching 'Lider Técnico' data found in Maturity file for the provided 'Matricula Lider'.")
    
    practices = df_filtered['PRACTICA'].unique()
    if len(practices) < 4:
        raise ValueError(f"Not enough practices for Maturity graphs; expected 4, got {len(practices)}.")
    
    maturity_images = []
    for practice in practices[:4]:
        df_practice = df_filtered[df_filtered['PRACTICA'] == practice]
        plt.figure(figsize=(10, 6))
        bars = plt.bar(df_practice['Lider Técnico'], df_practice[q], color='#000474')
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2, yval + 0.1, f'{yval:.2f}', ha='center', va='bottom', fontsize=10, color='black', zorder=3)
        plt.xlabel('Lider Técnico', fontsize=12)
        plt.ylabel(f'Nivel de Madurez ({q})', fontsize=12)
        plt.title(f'Niveles de Madurez para {practice}', fontsize=16)
        plt.xticks(rotation=45)
        plt.tight_layout()

        img_bytes = BytesIO()
        plt.savefig(img_bytes, format='png')
        img_bytes.seek(0)
        plt.close()
        maturity_images.append(img_bytes)
    return maturity_images

# --- Main Azure Function ---
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.function_name(name="generate_presentation")
@app.route(route="generate_presentation", methods=["POST"])
def generate_presentation(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Processing generate_presentation request.")
    auth_header = req.headers.get("Authorization")
    if not auth_header or auth_header != AUTH_TOKEN_EXPECTED:
        return func.HttpResponse("Unauthorized", status_code=401)
    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON in request body.", status_code=400)
    
    required_keys = ["q", "matricula_lider", "tmd_file", "users_file", "pases_file", "revisiones_file", "maturity_level_file"]
    missing_keys = [key for key in required_keys if key not in req_body]
    if missing_keys:
        return func.HttpResponse(f"Missing required keys: {missing_keys}", status_code=400)
    
    q = req_body["q"]
    matricula_lider = req_body["matricula_lider"]
    tmd_file = req_body["tmd_file"]
    users_file = req_body["users_file"]
    pases_file = req_body["pases_file"]
    revisiones_file = req_body["revisiones_file"]
    maturity_level_file = req_body["maturity_level_file"]
    
    with tempfile.TemporaryDirectory() as temp_dir:
        logging.info(f"Temporary directory created: {temp_dir}")
        try:
            tmd_file_path = download_file_from_share(temp_dir, tmd_file, INPUTS_DIRECTORY)
            users_file_path = download_file_from_share(temp_dir, users_file, INPUTS_DIRECTORY)
            pases_file_path = download_file_from_share(temp_dir, pases_file, INPUTS_DIRECTORY)
            revisiones_file_path = download_file_from_share(temp_dir, revisiones_file, INPUTS_DIRECTORY)
            maturity_file_path = download_file_from_share(temp_dir, maturity_level_file, INPUTS_DIRECTORY)
            
            template_filename = "base_template.pptx"
            template_file_path = download_file_from_share(temp_dir, template_filename, TEMPLATES_DIRECTORY)
            
            # Generate graphs with retry logic
            kpr_img = retry_call(lambda: generate_kpr_graph(tmd_file_path, users_file_path, matricula_lider, q))
            kr_img = retry_call(lambda: generate_kr_graph(pases_file_path, revisiones_file_path))
            maturity_imgs = retry_call(lambda: generate_maturity_graphs(maturity_file_path, users_file_path, matricula_lider, q))
            
            # Open the base presentation (template)
            prs = Presentation(template_file_path)
            
            # Create dynamic slides for each set of graphs
            create_dynamic_slide(prs, images=[kpr_img], title_text=f"KPR - {q}")
            create_dynamic_slide(prs, images=[kr_img], title_text="KR")
            create_dynamic_slide(prs, images=maturity_imgs, title_text="Niveles de Madurez")
            
            # Save presentation with date-based filename
            file_date = datetime.now().strftime("%m%d%y")
            output_filename = f"{file_date} Presentation.pptx"
            output_file_path = os.path.join(temp_dir, output_filename)
            prs.save(output_file_path)
            logging.info(f"Presentation saved at {output_file_path}")
            
            upload_file_to_share(output_file_path, output_filename, OUTPUTS_DIRECTORY)
            public_url = generate_file_url(output_filename, OUTPUTS_DIRECTORY)
            logging.info(f"Presentation generated successfully. Public URL: {public_url}")
            
            response_data = {"public_url": public_url}
            return func.HttpResponse(json.dumps(response_data), status_code=200, mimetype="application/json")
            
        except Exception as e:
            logging.error(f"Error in generate_presentation: {e}")
            return func.HttpResponse("An error occurred: " + str(e), status_code=500)
