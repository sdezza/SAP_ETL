# https://stackoverflow.com/questions/56494016/using-python3-sharepy-to-download-an-excel-file-from-a-shared-0365-commercial-sh
import sharepy
from sharepy import connect
from sharepy import SharePointSession
import os
import pandas as pd

server='https://group.sharepoint.com'
user='xxx@xxxx.com'
password='****'
data_file = "Reporting_ETL.xlsx"

# Copy/Paste file link from sharepoint
site = "https://group.sharepoint.com/sites/WORKGROUP/Shared%20Documents/5.%20MASTERDATA%20powerbi/Reporting_ETL.xlsx"


s = sharepy.connect(server,user,password)
# s.save()


def download_file():
  data_file = "Reporting_ETL.xlsx"
  #  Download file to same folder as python script.
  r = s.getfile(site,\
    filename = data_file)
  print("file downloaded")

  # Dataframe to send to script.py (easier with Pandas!)
  df_config = pd.read_excel(data_file, sheet_name="Config")
  df_config = df_config.astype(str)

  return df_config
  

def upload_file(filename):
  headers = {"accept": "application/json;odata=verbose",
  "content-type": "application/x-www-urlencoded; charset=UTF-8"}

  with open(filename, 'rb') as read_file:
    content = read_file.read()
    if filename == "Reporting_ETL.xlsx":
      p = s.post("https://group.sharepoint.com/sites/cnx_ops/_api/web/getfolderbyserverrelativeurl('/sites/WORKGROUP/Shared%20Documents/5.%20MASTERDATA%20powerbi/')/Files/add(url='"+filename+"',overwrite=true)", data=content, headers=headers)
    else:
      p = s.post("https://group.sharepoint.com/sites/cnx_ops/_api/web/getfolderbyserverrelativeurl('/sites/WORKGROUP/Shared%20Documents/5.%20MASTERDATA%20powerbi/DATASETS%20V2/')/Files/add(url='"+filename+"',overwrite=true)", data=content, headers=headers)
  
  print("file uploaded")
  # Delete the local uploaded file
  if os.path.exists(filename):
    os.remove(filename)
  else:
    print("The file does not exist")
