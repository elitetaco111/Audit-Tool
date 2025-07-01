from email.mime import base
import os
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
import concurrent.futures



def ensure_folder(folder_name):
    os.makedirs(folder_name, exist_ok=True)

def read_csv(csv_file):
    return pd.read_csv(csv_file, dtype=str)

def download_image(identifier, save_as, folder_name, base_url='https://media.rallyhouse.com/homepage/{}-1.jpg?tx=f_auto,c_fit,w_730,h_730'):
    image_path = os.path.join(folder_name, f"{save_as}.jpg")
    if os.path.exists(image_path):
        print(f"Skipped: {image_path} (already exists)")
        return False

    img_url = base_url.format(identifier)
    try:
        response = requests.get(img_url, headers={'User-Agent': 'Mozilla/5.0'}, stream=True)
        response.raise_for_status()
        image = Image.open(BytesIO(response.content))
        if image.mode == "P":
            image = image.convert("RGB")
        image.save(image_path)
        print(f"Downloaded: {image_path}")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to download {save_as}: {e}")
        return False

def download_images(csv_file, folder_name, item_col='Name', picture_id_col='Picture ID', max_workers=12, base_url='https://media.rallyhouse.com/homepage/{}-1.jpg?tx=f_auto,c_fit,w_730,h_730'):
    ensure_folder(folder_name)
    dataFile = read_csv(csv_file)
    if item_col not in dataFile.columns or picture_id_col not in dataFile.columns:
        raise ValueError(f"'{item_col}' or '{picture_id_col}' column not found in the CSV file.")
    tasks = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        for _, row in dataFile.iterrows():
            name = row[item_col]
            picture_id = row[picture_id_col]
            if name == picture_id:
                tasks.append(executor.submit(download_image, name, name, folder_name, base_url))
            else:
                tasks.append(executor.submit(download_image, picture_id, name, folder_name, base_url))
        concurrent.futures.wait(tasks)
    print('Download Complete')

if __name__ == "__main__":
    HALEY = 'https://media.rallyhouse.com/homepage/{}-1.jpg?tx=f_auto,c_fit,w_730,h_730'
    MY_FOREVER_LOVE = "osu"
    I_MISS_MY_GF = f'data_{MY_FOREVER_LOVE}.csv'
    I_LOVE_MY_GF_HALEY = f'{MY_FOREVER_LOVE}_images'
    download_images(I_MISS_MY_GF, I_LOVE_MY_GF_HALEY, base_url=HALEY)
