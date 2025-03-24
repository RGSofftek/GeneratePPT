"""
This Azure Function module generates a PowerPoint presentation by integrating several data visualizations.
It downloads input Excel files and a PowerPoint template from an Azure File Share, generates graphs based on 
the Excel data, inserts them into designated slides in the template, uploads the final presentation back to the share,
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
from pptx.util import Pt
from io import BytesIO
from azure.storage.fileshare import ShareServiceClient
import pandas as pd
import matplotlib.pyplot as plt
import tempfile

# Environment variables and storage structure definitions
STORAGE_ACCOUNT_NAME = os.environ.get("STORAGE_ACCOUNT_NAME")
SAS_TOKEN = os.environ.get("SAS_TOKEN")
AUTH_TOKEN_EXPECTED = os.environ.get("AUTH_TOKEN")  # Expected token for authentication
# Folder structure in the Azure File Share
FILE_SHARE_NAME = os.environ.get("FILE_SHARE_NAME") # Name of the Azure File Share
DIRECTORY_NAME = os.environ.get("DIRECTORY_NAME")  # Root directory for the function
TEMPLATES_DIRECTORY = f"{DIRECTORY_NAME}/templates"   # Folder for PowerPoint templates
INPUTS_DIRECTORY = f"{DIRECTORY_NAME}/inputs"           # Folder for Excel input files
OUTPUTS_DIRECTORY = f"{DIRECTORY_NAME}/outputs"         # Folder for generated presentations

# Predefined position for inserting KR and KPR graphs in the PPT slides
GRAPH_POSITION = (Pt(220), Pt(150), Pt(350), Pt(700))

# Initialize the Azure File Share client
service_url = f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net"
SERVICE_CLIENT = ShareServiceClient(account_url=service_url, credential=SAS_TOKEN)
SHARE_CLIENT = SERVICE_CLIENT.get_share_client(FILE_SHARE_NAME)

def download_file_from_share(temp_dir: str, filename: str, folder: str) -> str:
    """
    Downloads a file from the specified Azure File Share folder and saves it locally.
    
    Args:
        temp_dir (str): Local temporary directory to store the downloaded file.
        filename (str): Name of the file to download.
        folder (str): Folder in the Azure File Share where the file is stored.
    
    Returns:
        str: The full local path of the downloaded file.
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
    
    The function first attempts to delete an existing file (to simulate overwriting),
    then uploads the new file.
    
    Args:
        local_file_path (str): Full path of the local file.
        filename (str): Name to use for the file on the share.
        folder (str): Destination folder in the file share.
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
    
    Args:
        filename (str): Name of the file.
        folder (str): Folder in the file share where the file is stored.
    
    Returns:
        str: Public URL granting read access (SAS_TOKEN must be appropriately configured).
    """
    return f"https://{STORAGE_ACCOUNT_NAME}.file.core.windows.net/{FILE_SHARE_NAME}/{folder}/{filename}?{SAS_TOKEN}"

def retry_call(func_call, attempts: int = 2, delay: int = 1):
    """
    Executes a function with retry logic in case of transient errors.
    
    Args:
        func_call: A no-argument callable representing the function to execute.
        attempts (int): Maximum number of attempts (default is 2).
        delay (int): Delay in seconds between attempts (default is 1).
    
    Returns:
        The result of the callable if successful.
    
    Raises:
        Exception: The last exception encountered if all attempts fail.
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
    
    Args:
        slide: The pptx slide object.
        image (BytesIO): Image data as a BytesIO stream.
        position (tuple): Tuple specifying (left, top, width, height) using pptx.util units.
    """
    left, top, width, height = position
    slide.shapes.add_picture(image, left, top, width=width, height=height)

def generate_kpr_graph(tmd_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> BytesIO:
    """
    Generates the KPR graph using data from the TMD and Users Excel files.
    
    The function filters the TMD data for users managed by the provided 'matricula_lider' and plots a 
    bar chart of the values corresponding to the given quarter (or metric). It annotates each bar with its value.
    
    Args:
        tmd_file_path (str): Path to the TMD Excel file.
        users_file_path (str): Path to the Users Excel file.
        matricula_lider (str): The leader identifier used to filter users.
        q (str): The quarter or metric column name to plot.
    
    Returns:
        BytesIO: Stream containing the generated PNG image.
    
    Raises:
        ValueError: If required columns are missing or no matching data is found.
    """
    # Load and validate Users Excel file
    df_users = pd.read_excel(users_file_path)
    required_users_cols = {"Matricula Lider", "Matricula"}
    if not required_users_cols.issubset(set(df_users.columns)):
        raise ValueError(f"Users file is missing required columns: {required_users_cols - set(df_users.columns)}")
    
    # Load and validate TMD Excel file
    df_tmd = pd.read_excel(tmd_file_path)
    required_tmd_cols = {"Matricula LT", "Lider Técnico", q}
    if not required_tmd_cols.issubset(set(df_tmd.columns)):
        raise ValueError(f"TMD file is missing required columns: {required_tmd_cols - set(df_tmd.columns)}")
    
    # Filter users based on the provided 'Matricula Lider'
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    if not users_list:
        raise ValueError("No users found matching the provided 'Matricula Lider' in the Users file.")
    
    # Filter TMD data for matching user identifiers
    df_filtered = df_tmd[df_tmd["Matricula LT"].isin(users_list)]
    if df_filtered.empty:
        raise ValueError("No matching 'Lider Técnico' data found in TMD file for KPR graph.")
    
    # Create the KPR bar chart
    plt.figure(figsize=(10, 6))
    bars = plt.bar(df_filtered["Lider Técnico"], df_filtered[q], color='#000474')
    for bar in bars:
        yval = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,  # Horizontal center of the bar
            yval + 0.5,                        # Slightly above the bar
            f'{yval:.2f}',                     # Formatted value with 2 decimals
            ha='center',
            va='bottom',
            fontsize=10,
            color='black',
            zorder=3
        )
    
    # Set title and labels with appropriate font sizes
    plt.title(f'{q} Values by Selected Lider Técnico (KPR)', fontsize=16)
    plt.xlabel('Lider Técnico', fontsize=12)
    plt.ylabel(q, fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()

    # Save the figure to a BytesIO stream
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()
    return img_bytes


def generate_kr_graph(pases_file_path: str, revisiones_file_path: str) -> BytesIO:
    """
    Generates the KR graph using data from the Pases and Revisiones Excel files.
    
    Both Excel files are expected to contain specific quarter columns. The function sums the values 
    per quarter for each dataset and plots a side-by-side bar chart for comparison. Each bar is annotated with its value.
    
    Args:
        pases_file_path (str): Path to the Pases Excel file.
        revisiones_file_path (str): Path to the Revisiones Excel file.
    
    Returns:
        BytesIO: Stream containing the generated PNG image.
    
    Raises:
        ValueError: If required quarter columns are missing in either file.
    """
    # Define quarter categories
    quarters = ['Q1 01', 'Q1 02', 'Q1 03', 'Q2 01', 'Q2 02', 'Q2 03']
    
    # Load the Excel files
    df1 = pd.read_excel(pases_file_path)
    df2 = pd.read_excel(revisiones_file_path)
    
    # Validate required quarter columns exist
    missing_df1 = set(quarters) - set(df1.columns)
    missing_df2 = set(quarters) - set(df2.columns)
    if missing_df1:
        raise ValueError(f"Pases file is missing required columns: {missing_df1}")
    if missing_df2:
        raise ValueError(f"Revisiones file is missing required columns: {missing_df2}")
    
    # Sum values across the quarters for each dataset
    df1_grouped = df1[quarters].sum(axis=0)
    df2_grouped = df2[quarters].sum(axis=0)

    # Calculate positions for side-by-side bars
    bar_width = 0.4
    x = range(len(quarters))
    x1 = [pos - bar_width / 2 for pos in x]
    x2 = [pos + bar_width / 2 for pos in x]

    # Create the KR bar chart
    plt.figure(figsize=(10, 7))
    bars1 = plt.bar(x1, df1_grouped, color='#000474', width=bar_width, label='Dataset 1', edgecolor='black', zorder=2)
    bars2 = plt.bar(x2, df2_grouped, color='#f46a34', width=bar_width, label='Dataset 2', edgecolor='black', zorder=1)
    
    # Annotate bars for dataset 1
    for bar in bars1:
        yval = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            yval,
            round(yval, 2),
            va='bottom',
            ha='center',
            fontsize=10,
            color='black',
            zorder=3
        )
    # Annotate bars for dataset 2
    for bar in bars2:
        yval = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            yval,
            round(yval, 2),
            va='bottom',
            ha='center',
            fontsize=10,
            color='black',
            zorder=3
        )
    
    # Set title and axis labels
    plt.title('Pases Exitosos vs Reversiones x Q (KR)', fontsize=16)
    plt.xlabel('Trimestre', fontsize=12)
    plt.ylabel('Valor Total', fontsize=12)
    plt.xticks(ticks=x, labels=quarters, rotation=0)
    plt.grid(linestyle='--', linewidth=0.5)
    plt.legend()
    plt.tight_layout()

    # Save the figure to a BytesIO stream
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)
    plt.close()
    return img_bytes


def generate_maturity_graphs(maturity_file_path: str, users_file_path: str, matricula_lider: str, q: str) -> list:
    """
    Generates four maturity level graphs (one per practice) using data from the Maturity and Users Excel files.
    
    The function filters the maturity data for users under the specified 'matricula_lider' and creates 
    a bar chart for each of the first four practices found. Each graph is annotated with bar values.
    
    Args:
        maturity_file_path (str): Path to the Maturity Excel file.
        users_file_path (str): Path to the Users Excel file.
        matricula_lider (str): The leader identifier used to filter users.
        q (str): The quarter or metric column name to plot.
    
    Returns:
        list: A list of BytesIO streams, each containing a PNG image for a practice.
    
    Raises:
        ValueError: If required columns are missing, no matching users are found, 
                    or if there are fewer than four practices.
    """
    # Load and validate Users Excel file
    df_users = pd.read_excel(users_file_path)
    required_users_cols = {"Matricula Lider", "Matricula"}
    if not required_users_cols.issubset(set(df_users.columns)):
        raise ValueError(f"Users file is missing required columns: {required_users_cols - set(df_users.columns)}")
    
    # Load and validate Maturity Excel file
    df_maturity = pd.read_excel(maturity_file_path)
    required_maturity_cols = {"Matricula LT", "PRACTICA", "Lider Técnico", q}
    if not required_maturity_cols.issubset(set(df_maturity.columns)):
        raise ValueError(f"Maturity file is missing required columns: {required_maturity_cols - set(df_maturity.columns)}")
    
    # Filter users based on the provided 'Matricula Lider'
    users_list = df_users[df_users['Matricula Lider'] == matricula_lider]['Matricula'].tolist()
    if not users_list:
        raise ValueError("No users found matching the provided 'Matricula Lider' in the Users file.")
    
    # Filter maturity data for the selected users
    df_filtered = df_maturity[df_maturity["Matricula LT"].isin(users_list)]
    if df_filtered.empty:
        raise ValueError("No matching 'Lider Técnico' data found in Maturity file for the provided 'Matricula Lider'.")
    
    # Ensure there are at least four practices available
    practices = df_filtered['PRACTICA'].unique()
    if len(practices) < 4:
        raise ValueError(f"Not enough practices for Maturity graphs; expected 4, got {len(practices)}.")
    
    maturity_images = []
    # Generate graphs for each of the first four practices
    for practice in practices[:4]:
        df_practice = df_filtered[df_filtered['PRACTICA'] == practice]
        plt.figure(figsize=(10, 6))
        bars = plt.bar(df_practice['Lider Técnico'], df_practice[q], color='#000474')
        for bar in bars:
            yval = bar.get_height()
            plt.text(
                bar.get_x() + bar.get_width() / 2,
                yval + 0.1,
                f'{yval:.2f}',
                ha='center',
                va='bottom',
                fontsize=10,
                color='black',
                zorder=3
            )
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


app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.function_name(name="generate_presentation")
@app.route(route="generate_presentation", methods=["POST"])
def generate_presentation(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP-triggered Azure Function that creates a PowerPoint presentation with embedded graphs.
    
    Steps:
      1. Validates the request for proper JSON payload and authorization.
      2. Downloads required Excel files and the PowerPoint template from the Azure File Share.
      3. Generates the KPR, KR, and four Maturity graphs from the Excel files.
      4. Inserts the generated graphs into specific slides of the template.
      5. Saves the presentation locally, uploads it to the share, and returns a public URL.
    
    Returns:
        func.HttpResponse: On success, returns a JSON response with {"public_url": "<URL>"}.
                           On error, returns an appropriate error message.
    """
    logging.info("Processing generate_presentation request.")

    # Validate authorization
    auth_header = req.headers.get("Authorization")
    if not auth_header or auth_header != AUTH_TOKEN_EXPECTED:
        return func.HttpResponse("Unauthorized", status_code=401)
    
    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON in request body.", status_code=400)
    
    # Ensure all required keys are present in the request body
    required_keys = [
        "q", "matricula_lider", "tmd_file", "users_file", 
        "pases_file", "revisiones_file", "maturity_level_file"
    ]
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
    
    # Use a context manager to create a temporary directory that cleans up automatically
    with tempfile.TemporaryDirectory() as temp_dir:
        logging.info(f"Temporary directory created: {temp_dir}")
        try:
            # Download Excel input files and the PPT template from Azure File Share
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
            
            # Open the PowerPoint template
            prs = Presentation(template_file_path)
            if len(prs.slides) < 5:
                return func.HttpResponse("Template does not have enough slides for all graphs.", status_code=500)
            
            # Insert KR graph on slide 3 (index 2) and KPR graph on slide 4 (index 3)
            insert_picture_on_slide(prs.slides[2], kr_img, GRAPH_POSITION)
            insert_picture_on_slide(prs.slides[3], kpr_img, GRAPH_POSITION)
            
            # Insert the four Maturity graphs on slide 5 (index 4) at predefined positions
            maturity_positions = [
                (Pt(120), Pt(150), Pt(170), Pt(300)),
                (Pt(120), Pt(350), Pt(170), Pt(300)),
                (Pt(520), Pt(150), Pt(170), Pt(300)),
                (Pt(520), Pt(350), Pt(170), Pt(300))
            ]
            for img, pos in zip(maturity_imgs, maturity_positions):
                insert_picture_on_slide(prs.slides[4], img, pos)
            
            # Save the modified presentation with a date-based filename
            file_date = datetime.now().strftime("%m%d%y")
            output_filename = f"{file_date} Presentation.pptx"
            output_file_path = os.path.join(temp_dir, output_filename)
            prs.save(output_file_path)
            logging.info(f"Presentation saved at {output_file_path}")
            
            # Upload the presentation to the outputs folder and generate a public URL
            upload_file_to_share(output_file_path, output_filename, OUTPUTS_DIRECTORY)
            public_url = generate_file_url(output_filename, OUTPUTS_DIRECTORY)
            logging.info(f"Presentation generated successfully. Public URL: {public_url}")
            
            response_data = {"public_url": public_url}
            return func.HttpResponse(json.dumps(response_data), status_code=200, mimetype="application/json")
            
        except Exception as e:
            logging.error(f"Error in generate_presentation: {e}")
            return func.HttpResponse("An error occurred: " + str(e), status_code=500)
