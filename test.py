import requests
from io import BytesIO
import pandas as pd

url = "https://netorg18892072.sharepoint.com/:x:/s/allcompany/IQDilvFlWwHLRr0qX0NVe-SOAcrM_IsVLRb4Ea7JUQqn6g4?e=FU7VBf"

response = requests.get(url)

print("Status:", response.status_code)
print("Content-Type:", response.headers.get("content-type"))

try:
    df = pd.read_excel(BytesIO(response.content))
    print("Excel carregado com sucesso!")
    print(df.head())
except Exception as e:
    print("Erro ao abrir como Excel:")
    print(e)