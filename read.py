import win32com.client
import regex
import json
from json import dumps
from datetime import datetime
import pandas as pd
import numpy as np
import sys

# Parametros de fecha
fecha = sys.argv[1]
dateOut = datetime.strptime(fecha, '%d/%m/%Y')

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Validar parametro directorio
if len(sys.argv) > 2:
    dir = sys.argv[2]
    inbox = outlook.GetDefaultFolder(6).Folders.Item(dir)
else:
    inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items

idsVol = []
jsonVol = []
cont = 0

for message in messages:
    strSubject = message.subject
    pattern = regex.compile(r'\{(?:[^{}]|(?R))*\}')
    if 'OCI Event Notification' in strSubject:
        cont += 1
        #print (message.body)
        body = message.body
        jsonBody = pattern.findall(body)
        dicBody = json.loads(jsonBody[0])
        dateBackup = datetime.strptime(dicBody["eventTime"], '%Y-%m-%dT%H:%M:%S%z')

        if dateBackup.month == dateOut.month and dateBackup.day == dateOut.day:
            idsVol.append(dicBody["data"]["additionalDetails"]["volumeId"])

            if '.begin' in dicBody["eventType"]:
                jsonVol.append(
                    {
                        'volumeId': dicBody["data"]["additionalDetails"]["volumeId"],
                        'volumeName': '',
                        'compartmentName': dicBody["data"]["compartmentName"],
                        'eventTime': dateBackup.strftime("%Y-%m-%d %H:%M:%S"),
                        'backupState': dicBody["data"]["additionalDetails"]["backupState"]
                    }
                )
            else:
                jsonVol.append(
                    {
                        'volumeId': dicBody["data"]["additionalDetails"]["volumeId"],
                        'volumeName': dicBody["data"]["additionalDetails"]["volumeName"],
                        'compartmentName': dicBody["data"]["compartmentName"],
                        'eventTime': dateBackup.strftime("%Y-%m-%d %H:%M:%S"),
                        'backupState': dicBody["data"]["additionalDetails"]["backupState"]
                    }
                )

if cont <= 0:
    print("No hay mensajes sobre OCI con la fecha requerida en la bandeja de entrada, por favor seleccione algun otro directorio")
    quit()

# Read volumens.xlsx
dfV = pd.read_excel('volumes.xlsx')
# Conversion de list Dictionary a dataframe
df = pd.DataFrame(jsonVol, dtype="datetime64[ns]")
# Borrado de registro duplicados
dfR = df.drop_duplicates(subset=['volumeId', 'eventTime', 'backupState'], keep='first')
# Dataframe CREATE_PENDING
dfP = dfR.query("backupState=='CREATE_PENDING'")
# Dataframe AVAILABLE
dfA = dfR.query("backupState=='AVAILABLE'")
# Outer Join de dataframas CREATE_PENDING y AVAILABLE por volumeId
dfRR = pd.merge(dfP, dfA, how='outer', left_on='volumeId', right_on='volumeId')
# Remplazar los valores NaN por "X"
dfRR = dfRR.replace(np.nan, '')
#
dfRR['compartmentName_y'] = np.where(dfRR['compartmentName_y'] == '', dfRR['compartmentName_x'], dfRR['compartmentName_y'])
dfRR['compartmentName_x'] = np.where(dfRR['compartmentName_x'] == '', dfRR['compartmentName_y'], dfRR['compartmentName_x'])
dfRR = dfRR[['volumeId', 'volumeName_y', 'compartmentName_x', 'eventTime_x', 'backupState_x', 'eventTime_y', 'backupState_y']]
dfRR = dfRR.rename(columns={'volumeName_y':'volumeName', 'compartmentName_x':'compartmentName', 'eventTime_x':'eventTime_inicio', 'backupState_x':'backupState_inicio', 'eventTime_y':'eventTime_fin', 'backupState_y':'backupState_fin'})

dfRS = pd.merge(dfRR, dfV, how='outer', left_on='volumeId', right_on='volumeId')

# Creación de archivo de Excel
writer = pd.ExcelWriter(f'infVolumeBackup_{dateOut.day}-{dateOut.month}-{dateOut.year}.xlsx')
# Escritura de dataframe en archivo de excel
dfRS.to_excel(writer)
# Guardado de información en archivo de excel
writer.save()
