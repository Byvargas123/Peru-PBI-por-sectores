#Importación de librerías
import pandas as pd                          #Pip install pandas.
import os                                    #Pip install OS.
from openpyxl import load_workbook           #Pip install openpyxl.
import glob                                  #Abrir archivo con nombre.
import zipfile
                                             #pip install playwright, pip install xlml, playwright install chromium.
#Definiendo ruta de trabajo
ruta_proyectoconsolidado = "G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Base a procesar" # colocamos la carpeta deseada
os.chdir(ruta_proyectoconsolidado)
print(os.getcwd())

#DEBES ELIMINAR MANUALMENTE LOS ARCHIVOS DE LA CARPETA BASE A PROCESAR

#### Scrapeando la página de INEI ###
from playwright.sync_api import sync_playwright
from os import path, mkdir

download_path = "G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Base a procesar"
if not path.exists(download_path):
    mkdir(download_path)

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    page = browser.new_page()
    link = "https://www.gob.pe/inei"
    page.goto(link)

    page.select_option("//select[contains(@name,'filters')]", value="excel")
        
    with page.expect_download() as download_info:
      
        page.click('//*[@id="download-resumen_5-mensual-report"]/button')

    download = download_info.value
    download_path_file = path.join(download_path, download.suggested_filename)
    download.save_as(download_path_file)

    browser.close()


#### Descomprimiendo archivo zip 

ruta_carpeta = "G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Base a procesar"

ruta_destino = "G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Base a procesar"

archivo_zip = None                                                 # Buscar el archivo zip en la carpeta
for archivo in os.listdir(ruta_carpeta):
    if archivo.endswith('.zip'):
        archivo_zip = os.path.join(ruta_carpeta, archivo)
        break

if archivo_zip:                                                    # Verifica si la ruta de destino existe, si no, la crea
    
    os.makedirs(ruta_destino, exist_ok=True)

                                                                   # Abre y extrae el archivo zip
    with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
        zip_ref.extractall(ruta_destino)

    print(f'Archivo descomprimido en: {ruta_destino}')
else:
    print('No se encontró ningún archivo ZIP en la carpeta.')
                  

##EXCEL VAB
df_VAB=pd.read_excel(glob.glob('VA*.xlsx')[0], sheet_name=0, usecols='A:L', skiprows=3)
df_VAB = df_VAB.dropna(subset=['PBI'])                                                                  #Elimino registros en blanco
df_VAB.columns = [str(col).strip() for col in df_VAB.columns]                                           #Convertimos a texto y quitamos espacios a los nombres de las columnas
df_VAB['Año y Mes'] = df_VAB['Año y Mes'].astype(int).astype(str)                                       #Las fechas las volvemos a texto sin decimal
print(df_VAB)

##EXCEL AGROPECUARIO
df_agro = pd.read_excel(glob.glob('1*.xlsx')[0], sheet_name=2, usecols='A:CJ', skiprows=4)
df_agro = df_agro.dropna(subset=['Agropecuario Total'])                                                 #Elimino registros en blanco
df_agro.columns = [str(col).strip() for col in df_agro.columns]                                         #Convertimos a texto y quitamos espacios a los nombres de las columnas
df_agro['Periodo'] = df_agro['Periodo'].astype(int).astype(str)                                         #Las fechas las volvemos a texto sin decimal
print(df_agro)

colum_agrodiv = [col for col in df_agro.columns if col not in ['Agropecuario Total', 'Periodo']]

for col in colum_agrodiv:
    df_agro[col] = (df_agro[col] / df_agro['Agropecuario Total'])

df_agro.drop(columns=['Agropecuario Total', 'Agrícola', 'Pecuario'], inplace=True)                      #cuidado no juntar con código inmediatamente antertior
print(df_agro)

'''comprobando suma 100%
columnas_select= [col for col in df_agro.columns if col != 'Periodo']
df_agro['Total'] = df_agro[columnas_select].sum(axis=1)
valores_unicos = df_agro['Total'].unique()
print(valores_unicos)

df_agro = df_agro.drop(columns=['Total']) 
del columnas_select
del valores_unicos'''


df_agro = pd.merge(df_agro, df_VAB[['Año y Mes','Agricultura, ganadería, caza y silvicultura']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')
df_agro = df_agro.drop('Año y Mes', axis=1)

colum_agromult = [col for col in df_agro.columns if col not in ['Agricultura, ganadería, caza y silvicultura', 'Periodo']]
for col in colum_agromult:
    df_agro[col] = df_agro[col] * df_agro['Agricultura, ganadería, caza y silvicultura']

print(df_agro)

'''comprobando suma suma total VAB
columnas_select1= [col for col in df_agro.columns if col not in ['Agricultura, ganadería, caza y silvicultura', 'Periodo']]
df_agro['Total'] = df_agro[columnas_select1].sum(axis=1)
df_agro['Dif'] = df_agro['Agricultura, ganadería, caza y silvicultura ']- df_agro['Total']
valores_unicos = df_agro['Dif'].unique()
print(valores_unicos)

df_agro = df_agro.drop(columns=['Total', 'Dif'])'''

##EXCEL PESCA
df_pesca = pd.read_excel(glob.glob('2*.xls')[0], sheet_name=1, usecols='A:DQ', skiprows=7)
print(df_pesca)

df_pesca.columns=['Periodo','Total_pesca','Total_maritimo','Total_marítimo_chd','Total_m_chd_cong.','Perico_m_chd_cong.','Caballa_m_chd_cong.',
 'Jurel_m_chd_cong.','Merluza_m_chd_cong.','Anguila_m_chd_cong.','Anchoveta_m_chd_cong.','Pejerrey_m_chd_cong.','Atún_m_chd_cong.',
 'Bonito_m_chd_cong.','Machete_m_chd_cong.','Sierra_m_chd_cong.','Liza_m_chd_cong.','Volador_m_chd_cong.','Tollo_m_chd_cong.','Tiburón_m_chd_cong.',
 'Otros Pescados_m_chd_cong.','Langostino_m_chd_cong.','Pota_m_chd_cong.','Concha de Abanico_m_chd_cong.','Calamar_m_chd_cong.','Abalón_m_chd_cong.',
 'Pulpo_m_chd_cong.','Caracol_m_chd_cong.','Concha Navaja_m_chd_cong.','Otros mariscos_m_chd_cong.','Erizo_m_chd_cong.','Pepino de Mar_m_chd_cong.',
 'Total_m_chd_enla.','Jurel_m_chd_enla.','Anchoveta_m_chd_enla.','Caballa_m_chd_enla.','Atún_m_chd_enla.','Machete_m_chd_enla.','Bonito_m_chd_enla.',
 'Sardina_m_chd_enla.','Otros Pescados_m_chd_enla.','Caracol_m_chd_enla.','Pota_m_chd_enla.','Concha Navaja_m_chd_enla.','Almejas_m_chd_enla.','Abalón_m_chd_enla.',
 'Otros Mariscos_m_chd_enla.','Total_m_chd_fresc.','Jurel_m_chd_fresc.','Perico_m_chd_fresc.','Corvina_m_chd_fresc.','Bonito_m_chd_fresc.','Liza_m_chd_fresc.',
 'Caballa_m_chd_fresc.','Pejerrey_m_chd_fresc.','Tollo_m_chd_fresc.','Lorna_m_chd_fresc.','Tiburón_m_chd_fresc.','Chiri_m_chd_fresc.','Cabrilla_m_chd_fresc.',
 'Cojinova_m_chd_fresc.','Coco o Suco_m_chd_fresc.','Ayanque (Cachema)_m_chd_fresc.','Machete_m_chd_fresc.','Merluza_m_chd_fresc.','Sardina_m_chd_fresc.',
 'Lenguado_m_chd_fresc.','Raya_m_chd_fresc.','Ojo de Uva_m_chd_fresc.','Pardo_m_chd_fresc.','Anchoveta_m_chd_fresc.','Atún_m_chd_fresc.','Otros Pescados_m_chd_fresc.',
 'Calamar_m_chd_fresc.','Pota_m_chd_fresc.','Concha de Abanico_m_chd_fresc.','Langostino_m_chd_fresc.','Choro_m_chd_fresc.','Caracol_m_chd_fresc.','Pulpo_m_chd_fresc.',
 'Almejas_m_chd_fresc.','Concha Negra_m_chd_fresc.','Cangrejo_m_chd_fresc.','Palabritas_m_chd_fresc.','Otros Mariscos_m_chd_fresc.','Otras Especies_m_chd_fresc.',
 'Total_m_chd_curad.','Caballa_m_chd_curad.','Jurel_m_chd_curad.','Liza_m_chd_curad.','Anchoveta_m_chd_curad.','Tollo_m_chd_curad.','Raya_m_chd_curad.','Perico_m_chd_curad.',
 'Cabrilla_m_chd_curad.','Merluza_m_chd_curad.','Otros Pescados_m_chd_curad.','Pota_m_chd_curad.','Otras Especies_m_chd_curad.','Total_marítimo_chi','Anchoveta_m_chi',
 'Otras Especies_m_chi','Total_continental','Total_cont_chd_cong.','Trucha_cont_chd_cong.','Total_cont_chd_curad.','Boquichico_cont_chd_curad.','Llambina_cont_chd_curad.',
 'Zúngaro_cont_chd_curad.','Trucha_cont_chd_curad.','Otros pescados_cont_chd_curad.','Total_cont_chd_fresc.','Trucha_cont_chd_fresc.','Boquichico_cont_chd_fresc.',
 'Zúngaro_cont_chd_fresc.','Llambina_cont_chd_fresc.','Tilapia_cont_chd_fresc.','Palometa_cont_chd_fresc.','Ractacara_cont_chd_fresc.','Pejerrey_cont_chd_fresc.',
 'Otros pescados_cont_chd_fresc.']

df_pesca.columns = [str(col).strip() for col in df_pesca.columns]  
df_pesca['Periodo'] = df_pesca['Periodo'].astype(int).astype(str) 
print(df_pesca)

colum_pescahdiv = [col for col in df_pesca.columns if col not in ['Periodo','Total_pesca']]

for col in colum_pescahdiv:
    df_pesca[col] = (df_pesca[col] / df_pesca['Total_pesca'])

df_pesca.drop(columns=['Total_pesca','Total_maritimo','Total_marítimo_chd','Total_m_chd_cong.','Total_m_chd_enla.','Total_m_chd_fresc.','Total_m_chd_curad.','Total_marítimo_chi', 
                       'Total_continental','Total_cont_chd_cong.','Total_cont_chd_curad.','Total_cont_chd_fresc.'], inplace=True)     #cuidado no juntar con código inmediatamente antertior                                                  

df_pesca = pd.merge(df_pesca, df_VAB[['Año y Mes','Pesca y acuicultura']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')

df_pesca = df_pesca.drop('Año y Mes', axis=1)

colum_pescadromult = [col for col in df_pesca.columns if col not in ['Pesca y acuicultura', 'Periodo']]
for col in colum_pescadromult:
    df_pesca[col] = df_pesca[col] * df_pesca['Pesca y acuicultura']

print(df_pesca)

##EXCEL MINERÍA
df_minehidro = pd.concat([pd.read_excel(glob.glob('3*.xlsx')[0], sheet_name=0, usecols='A:A',skiprows=4),
                          pd.read_excel(glob.glob('3*.xlsx')[0], sheet_name=0, usecols='N:AA',skiprows=4)], 
                          axis=1)
df_minehidro.columns= ['Periodo','Total_Mine_Hidro','Mineria metalica','Cobre','Oro','Zinc','Plata','Hierro','Plomo','Estaño','Molibdeno','Hidrocarburos','Petroleo crudo', 'Líquido de gas natural', 'Gas natural']
df_minehidro.columns = [str(col).strip() for col in df_minehidro.columns]                                                  #Convertimos a texto y quitamos espacios a los nombres de las columnas
df_minehidro['Periodo'] = df_minehidro['Periodo'].astype(int).astype(str) 
print(df_minehidro.columns)

colum_minhdiv = [col for col in df_minehidro.columns if col not in ['Total_Mine_Hidro', 'Periodo']]

for col in colum_minhdiv:
    df_minehidro[col] = (df_minehidro[col] / df_minehidro['Total_Mine_Hidro'])

df_minehidro.drop(columns=['Total_Mine_Hidro','Mineria metalica','Hidrocarburos'], inplace=True)                            #cuidado no juntar con código inmediatamente antertior

df_minehidro = pd.merge(df_minehidro, df_VAB[['Año y Mes','Extraccion de petróleo, gas, minerales y servicios conexos']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')

df_minehidro = df_minehidro.drop('Año y Mes', axis=1)

colum_minehidromult = [col for col in df_minehidro.columns if col not in ['Extraccion de petróleo, gas, minerales y servicios conexos', 'Periodo']]
for col in colum_minehidromult:
    df_minehidro[col] = df_minehidro[col] * df_minehidro['Extraccion de petróleo, gas, minerales y servicios conexos']

print(df_minehidro)


'''comprobando suma suma total VAB
columnas_select1= [col for col in df_minehidro.columns if col not in ['Extraccion de petróleo, gas, minerales y servicios conexos', 'Periodo']]
df_minehidro['Total'] = df_minehidro[columnas_select1].sum(axis=1)
df_minehidro['Dif'] = df_minehidro['Extraccion de petróleo, gas, minerales y servicios conexos']- df_minehidro['Total']
valores_unicos = df_minehidro['Dif'].unique()
print(valores_unicos)

df_minehidro = df_minehidro.drop(columns=['Total', 'Dif'])'''


##EXCEL MANUFACTURA
df_manufac = pd.read_excel(glob.glob('4*.xlsx')[0], sheet_name=1, usecols='A:FM', skiprows=2)
print(df_manufac)

df_manufac = df_manufac.rename(columns={'CIIU': 'Periodo', 'TOTAL': 'Total_Manufac'})                         #Renombrando periodo

df_manufac.columns = [str(col).strip() for col in df_manufac.columns]                                         #Aplicando texto y quitando espacio a nombre de campos

df_manufac.drop(df_manufac.columns[[2,3,4, 5, 7, 9, 11, 13, 15, 18, 24, 26, 27, 32, 33, 37, 42, 43, 45, 47, 
           48, 51, 53, 54, 56, 60, 61, 65, 66, 69, 70, 72, 73, 77, 82, 84, 85, 
           87, 88, 91, 93, 94, 96, 104, 105, 107, 109, 112, 113, 116, 120, 121, 
           123, 124, 126, 128, 130, 132, 133, 138, 142, 143, 145, 147, 149, 150, 
           152, 155, 156, 158, 159, 162, 164, 165]], axis=1, inplace=True)                                    #Eliminamos columnas de CIIU 2 dig y 3 dig

prim_fila_pondera = df_manufac.iloc[0]                                                                        #Extraemos primera fila, que posee los ponderadores 2007

excluye_column = 'Periodo'

for col in df_manufac.columns:
    if col != excluye_column:
        df_manufac[col] = df_manufac[col] * prim_fila_pondera[col]                                            #Multiplicamos cada fila por el ponderador2007

df_manufac.drop(index=[0, 1], inplace=True)                                                                   #Eliminamos 2 primeros registros

df_manufac['Periodo'] = df_manufac['Periodo'].astype(int).astype(str) 

colum_manufachdiv = [col for col in df_manufac.columns if col not in ['Total_Manufac', 'Periodo']]

for col in colum_manufachdiv:
    df_manufac[col] = (df_manufac[col] / df_manufac['Total_Manufac'])

df_manufac.drop(columns=['Total_Manufac'], inplace=True)                                                        #cuidado no juntar con código inmediatamente antertior

df_manufac = pd.merge(df_manufac, df_VAB[['Año y Mes','Manufactura']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')

df_manufac = df_manufac.drop('Año y Mes', axis=1)

colum_manufactmult = [col for col in df_manufac.columns if col not in ['Manufactura', 'Periodo']]
for col in colum_manufactmult:
    df_manufac[col] = df_manufac[col] * df_manufac['Manufactura']

correla_ciiu4=pd.read_excel('G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Script y correlacionadores/Correla_ciiu4.xlsx', sheet_name='Correla', keep_default_na=False, usecols='A:B')                             #Importamos correlcionador CIIU4

correla_ciiu4['CIIU4-4dig'] = correla_ciiu4['CIIU4-4dig'].astype(str)                                                                     #Convertimos a texto la columna CIIU

mapeo_dict = dict(zip(correla_ciiu4['CIIU4-4dig'], correla_ciiu4['CIIU4-4dig-Descrición']))                                               #Lo volvemos un diccionario

df_manufac.columns = [mapeo_dict.get(col, col) for col in df_manufac.columns]                                                             #Reemplazamos el diccionario para reemplazar nombre de columnas

print(df_manufac)


##EXCEL ELECTRICIDAD Y AGUA
df_electragua = pd.read_excel(glob.glob('5*.xls')[0], sheet_name=0, usecols='H:L', skiprows=6)
print(df_electragua)
df_electragua.columns= ['Periodo','Total_Electragua','Electricidad de servicio público','Gas','Agua']

prim_fila_pondera = df_electragua.iloc[0]                                                                           #Extraemos primera fila, que posee los ponderadores 2007
print(prim_fila_pondera)
excluye_column = 'Periodo'

for col in df_electragua.columns:
    if col != excluye_column:
        df_electragua[col] = df_electragua[col] * prim_fila_pondera[col]                                             #Multiplicamos cada fila por el ponderador207

df_electragua.drop(index=[0,1,2,3,4,5,6,7,8,9,10,11,12], inplace=True)                                               #Eliminamos registros de ponderación y 2011

df_electragua['Periodo'] = df_electragua['Periodo'].astype(int).astype(str) 

colum_electraguahdiv = [col for col in df_electragua.columns if col not in ['Total_Electragua', 'Periodo']]

for col in colum_electraguahdiv:
    df_electragua[col] = (df_electragua[col] / df_electragua['Total_Electragua'])

df_electragua.drop(columns=['Total_Electragua'], inplace=True)                                                        #cuidado no juntar con código inmediatamente antertior

df_electragua = pd.merge(df_electragua, df_VAB[['Año y Mes','Electricidad, gas, suministro de agua, alcantarillado y gestión de desechos y saneamiento']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')

df_electragua = df_electragua.drop('Año y Mes', axis=1)

colum_electraguamult = [col for col in df_electragua.columns if col not in ['Electricidad, gas, suministro de agua, alcantarillado y gestión de desechos y saneamiento', 'Periodo']]
for col in colum_electraguamult:
    df_electragua[col] = df_electragua[col] * df_electragua['Electricidad, gas, suministro de agua, alcantarillado y gestión de desechos y saneamiento']

print(df_electragua)


##EXCEL CONSTRUCCIÓN
df_construcc = pd.read_excel(glob.glob('6*.xlsx')[0], sheet_name=0, usecols='A:E', skiprows=4)

df_construcc.columns = [str(col).strip() for col in df_construcc.columns]  

df_construcc = df_construcc.dropna(subset=['CONCRETO'])  

df_construcc.columns= ['Periodo','Total_Construcción','Consumo interno de cemento','Vivienda no concreto','Avance de físico obras']

df_construcc['Periodo'] = df_construcc['Periodo'].astype(int).astype(str)

nuevo_registro = {'Periodo': 'XXXXXX', 'Total_Construcción': 1.00,'Consumo interno de cemento': 0.7395,'Vivienda no concreto': 0.0276,'Avance de físico obras': 0.2329} #agregamos ponderadores 2007

df_construcc.loc[len(df_construcc)] = nuevo_registro

ult_fila_pondera = df_construcc.iloc[-1]                                                                           #Extraemos primera fila, que posee los ponderadores 2007

excluye_column = 'Periodo'

for col in df_construcc.columns:
    if col != excluye_column:
        df_construcc[col] = df_construcc[col] * ult_fila_pondera[col]  


df_construcc.drop(index=df_construcc.index[-1], inplace=True)

colum_construcchdiv = [col for col in df_construcc.columns if col not in ['Total_Construcción', 'Periodo']]

for col in colum_construcchdiv:
    df_construcc[col] = (df_construcc[col] / df_construcc['Total_Construcción'])

df_construcc.drop(columns=['Total_Construcción'], inplace=True)                                                        #cuidado no juntar con código inmediatamente antertior

df_construcc = pd.merge(df_construcc, df_VAB[['Año y Mes','Construcción']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')

df_construcc = df_construcc.drop('Año y Mes', axis=1)

colum_construccmult = [col for col in df_construcc.columns if col not in ['Construcción', 'Periodo']]
for col in colum_construccmult:
    df_construcc[col] = df_construcc[col] * df_construcc['Construcción']

print(df_construcc)

##EXCEL COMERCIO
df_comercio = df_VAB[['Año y Mes', 'Comercio y mantenimiento y reparación de vehículos automotores y motocicletas']]

df_comercio = df_comercio.rename(columns={'Año y Mes': 'Periodo'})

df_comercio['Periodo'] = df_comercio['Periodo'].astype(int).astype(str)

print(df_comercio)


##EXCEL IMPUESTOS
df_impuestos = df_VAB[['Año y Mes','Derechos de Importación y Otros Impuestos a los productos (*)']]

df_impuestos = df_impuestos.rename(columns={'Año y Mes': 'Periodo', 'Derechos de Importación y Otros Impuestos a los productos (*)': 'Derechos de importación e impuestos'})

print(df_impuestos)

##PARA SECTORES DE SERVICIOS
df_VBP=pd.read_excel(glob.glob('Ind*.xlsx')[0], sheet_name=1, usecols='A:L', skiprows=3)
df_VBP = df_VBP.iloc[2:]
df_VBP.columns = [str(col).strip() for col in df_VBP.columns]                                                 #Convertimos a texto y quitamos espacios a los nombres de las columnas
df_VBP = df_VBP.dropna(subset=['Año y Mes'])                                                                  #Elimino registros en blanco   
df_VBP['Año y Mes'] = df_VBP['Año y Mes'].astype(int).astype(str)                                             #Las fechas las volvemos a texto sin decimal
print(df_VBP)

##EXCEL TRANSPORTE
df_Transp = pd.read_excel(glob.glob('8*.xlsx')[0], sheet_name=0, skiprows=2)

df_Transp = df_Transp.T                                                     #Trasponemos el dataframe
df_Transp.columns = df_Transp.iloc[1]                                       #Asignar 2da fila como encabezado  
df_Transp = df_Transp[2:]                                                   #Eliminar la fila de encabezados

df_Transp = df_Transp.reset_index(drop=True)                                #Reseteamos los índices

df_Transp = df_Transp.iloc[:, [0, 4, 5,7,8,9,11,12,14,15,18,19,21,22]]      #Nos quedamos con las columnas de menor jerarquía

df_Transp.columns = [str(col).strip() for col in df_Transp.columns]         #Quitamos los espacios al nombre de los campos

df_Transp = df_Transp.dropna(subset=['Transporte y Almacenamiento'])        #Elimimos último registro con NA

for col in df_Transp.columns:
    df_Transp[col] = pd.to_numeric(df_Transp[col], errors='coerce')         #Convertimos todas las columnas a float(extrañamente todos son objetos)

df_Transp.iloc[0] = df_Transp.iloc[0] / 100                              

df_Transp.iloc[0] = df_Transp.iloc[0]* 0.04969 

prim_fila_pondera1 = df_Transp.iloc[0]                                      #Extraemos primera fila, que posee los ponderadores 2007

for col in df_Transp.columns:
    df_Transp[col] = df_Transp[col] * prim_fila_pondera1[col]

df_Transp = df_Transp.drop(index=0)                                         #Eliminamos la fila de ponderadores

num_registros = len(df_Transp)                                              #Contamos el número de registros para posteriormente crear registros de fecha


df_Transp['Periodo'] = pd.date_range(start='2012-01-01', periods=num_registros, freq='ME').strftime('%Y%m')      #Creamos los registros de fecha       


df_Transp = df_Transp[['Periodo'] + [col for col in df_Transp.columns if col != 'Periodo']]                    #Reordenamos columnas    

df_Transp = pd.merge(df_Transp, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Transp = df_Transp.drop('Año y Mes', axis=1)

colum_Transphdiv = [col for col in df_Transp.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Transphdiv:
    df_Transp[col] = (df_Transp[col] / df_Transp['Índice Global'])                                           #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Transp.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Transp = pd.merge(df_Transp, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Transp = df_Transp.drop('Año y Mes', axis=1)                                                               

colum_Transphdiv2 = [col for col in df_Transp.columns if col not in ['PBI', 'Periodo']]

for col in colum_Transphdiv2:
    df_Transp[col] = df_Transp[col] * df_Transp['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Transp = df_Transp.drop('PBI', axis=1) 

print(df_Transp)


##EXCEL RESTAURANTE
df_Alojrest = pd.read_excel(glob.glob('9*.xls')[0], sheet_name=0, usecols='A:D', skiprows=10)

df_Alojrest.columns= ['Periodo','Total_Alojrest','Alojamiento','Restaurantes']

df_Alojrest.columns = [str(col).strip() for col in df_Alojrest.columns]  

df_Alojrest['Periodo'] = df_Alojrest['Periodo'].astype(int).astype(str)

nuevo_registro = {'Periodo': 'XXXXXX', 'Total_Alojrest': 1.00*0.0285993124653965,'Alojamiento': 0.1360*0.0285993124653965 ,'Restaurantes': 0.8640*0.0285993124653965}         #agregamos ponderadores 2007

df_Alojrest.loc[len(df_Alojrest)] = nuevo_registro

ult_fila_pondera = df_Alojrest.iloc[-1]                                                                           #Extraemos últimA fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                        #Excluímos la columna periodo

for col in df_Alojrest.columns:
    if col != excluye_column:
        df_Alojrest[col] = df_Alojrest[col] * ult_fila_pondera[col]                                               #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Alojrest.drop(index=df_Alojrest.index[-1], inplace=True)                                                       #Eliminamos la fila de ponderadores

df_Alojrest = pd.merge(df_Alojrest, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Alojrest = df_Alojrest.drop('Año y Mes', axis=1)

colum_Alojaresthdiv = [col for col in df_Alojrest.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Alojaresthdiv:
    df_Alojrest[col] = (df_Alojrest[col] / df_Alojrest['Índice Global'])                                          #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Alojrest.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Alojrest = pd.merge(df_Alojrest, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Alojrest = df_Alojrest.drop('Año y Mes', axis=1)                                                               

colum_Alojaresthdiv2 = [col for col in df_Alojrest.columns if col not in ['PBI', 'Periodo']]

for col in colum_Alojaresthdiv2:
    df_Alojrest[col] = df_Alojrest[col] * df_Alojrest['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Alojrest = df_Alojrest.drop('PBI', axis=1) 

print(df_Alojrest)

###Lo que stoy haciendo es: en el excel de restaurante, aplico el peso interno y luego el peso de restaurantes en pbi vbp (obteniendo el indice vbp estandarizado) y luego
# divido por el indice vbp pbi para sacar particpación en cada fila. Esa participáción la multiplico por el PBI VAB.
# Luego de hacer esto con cada sector servicios, el valor de otros servicios será la diferencia de sumatoria de los servioios scrtpt con el vab otros servicios.
# Ojo hemos notado que probablemente al final haya un residuio entre sumatoria de todo los sectores y el vab total. 


##EXCEL TELECOMUNICACIONES
df_Telecom = pd.read_excel(glob.glob('10*.xlsx')[0], sheet_name=0, usecols='A:D', skiprows=3)

df_Telecom.columns= ['Periodo','Total_Telecomyotro','Telecomunicaciones','Otros servicios de información']

df_Telecom.columns = [str(col).strip() for col in df_Telecom.columns]  

primera_fila_pondera2 = df_Telecom.iloc[0]                                                                                #Extraemos primera fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                                 #Excluímos la columna periodo

for col in df_Telecom.columns:
    if col != excluye_column:
        df_Telecom[col] = df_Telecom[col] * (primera_fila_pondera2[col]/100)                                               #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Telecom.drop(index=df_Telecom.index[0], inplace=True)                                                                    #Eliminamos fila de ponderadores

df_Telecom['Periodo'] = df_Telecom['Periodo'].astype(int).astype(str)

df_Telecom = df_Telecom.reset_index(drop=True)                                                                              #Resetamos índices porque por alguna razón el registro que se creará está reempazando al último                                                 

nuevo_registro = {'Periodo': "XXXXXX", 'Total_Telecomyotro': 0.0266411838857904, 'Telecomunicaciones':  0.0266411838857904 , 'Otros servicios de información': 0.0266411838857904}

df_Telecom.loc[len(df_Telecom)] = nuevo_registro

ult_registro_pondera = df_Telecom.iloc[-1]                                                                                   #Extraemos últimA fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                                   #Excluímos la columna periodo

for col in df_Telecom.columns:
    if col != excluye_column:
        df_Telecom[col] = df_Telecom[col] * ult_registro_pondera[col]                                                          #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Telecom.drop(index=df_Telecom.index[-1], inplace=True)                                                                     #Eliminamos la fila de ponderadores


df_Telecom = pd.merge(df_Telecom, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Telecom = df_Telecom.drop('Año y Mes', axis=1)

colum_Telecomdiv = [col for col in df_Telecom.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Telecomdiv:
    df_Telecom[col] = (df_Telecom[col] / df_Telecom['Índice Global'])                                            #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Telecom.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Telecom = pd.merge(df_Telecom, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Telecom = df_Telecom.drop('Año y Mes', axis=1)                                                               

colum_Telecomdiv2 = [col for col in df_Telecom.columns if col not in ['PBI', 'Periodo']]

for col in colum_Telecomdiv2:
    df_Telecom[col] = df_Telecom[col] * df_Telecom['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Telecom = df_Telecom.drop('PBI', axis=1) 

print(df_Telecom)


##EXCEL FINANCIERO

df_Financi = pd.read_excel(glob.glob('11*.xls')[0], sheet_name=0, usecols='A:B', skiprows=3)

df_Financi.columns= ['Periodo','Servicios Financieros']

num_registros = len(df_Financi)                                                                                  #Contamos el número de registros para posteriormente crear registros de fecha

df_Financi['Periodo'] = pd.date_range(start='2012-01-01', periods=num_registros, freq='ME').strftime('%Y%m')      #Creamos los registros de fecha       

df_Financi = df_Financi[['Periodo'] + [col for col in df_Financi.columns if col != 'Periodo']] 

nuevo_registro = {'Periodo': "XXXXXX", 'Servicios Financieros': 0.0321527215172056}

df_Financi.loc[len(df_Financi)] = nuevo_registro

ult_registro_pondera = df_Financi.iloc[-1]                                                                                   #Extraemos últimA fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                                   #Excluímos la columna periodo

for col in df_Financi.columns:
    if col != excluye_column:
        df_Financi[col] = df_Financi[col] * ult_registro_pondera[col]                                                          #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Financi.drop(index=df_Financi.index[-1], inplace=True) 


df_Financi = pd.merge(df_Financi, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Financi = df_Financi.drop('Año y Mes', axis=1)

colum_Financithdiv = [col for col in df_Financi.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Financithdiv:
    df_Financi[col] = (df_Financi[col] / df_Financi['Índice Global'])                                          #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Financi.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Financi = pd.merge(df_Financi, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Financi = df_Financi.drop('Año y Mes', axis=1)                                                               

colum_Financithdiv2 = [col for col in df_Financi.columns if col not in ['PBI', 'Periodo']]

for col in colum_Financithdiv2:
    df_Financi[col] = df_Financi[col] * df_Financi['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Financi = df_Financi.drop('PBI', axis=1) 

print(df_Financi)

##EXCEL SERVICIOS PRESTADOS A EMPRESAS
df_Servicempresas = pd.read_excel(glob.glob('12*.xls')[0], sheet_name=0, usecols='A:B', skiprows=4)

df_Servicempresas.columns= ['Periodo','Servicios prestados a empresas']

df_Servicempresas = df_Servicempresas.dropna(subset=['Servicios prestados a empresas'])   

df_Servicempresas['Periodo'] = df_Servicempresas['Periodo'].astype(int).astype(str)

nuevo_registro = {'Periodo': "XXXXXX", 'Servicios prestados a empresas': 0.0424000525504156}

df_Servicempresas.loc[len(df_Servicempresas)] = nuevo_registro

ult_registro_pondera = df_Servicempresas.iloc[-1]                                                                                   #Extraemos últimA fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                                   #Excluímos la columna periodo

for col in df_Servicempresas.columns:
    if col != excluye_column:
        df_Servicempresas[col] = df_Servicempresas[col] * ult_registro_pondera[col]                                                          #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Servicempresas.drop(index=df_Servicempresas.index[-1], inplace=True) 

df_Servicempresas = pd.merge(df_Servicempresas, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Servicempresas = df_Servicempresas.drop('Año y Mes', axis=1)

colum_Servicemprdiv = [col for col in df_Servicempresas.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Servicemprdiv:
    df_Servicempresas[col] = (df_Servicempresas[col] / df_Servicempresas['Índice Global'])                                          #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Servicempresas.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Servicempresas = pd.merge(df_Servicempresas, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Servicempresas = df_Servicempresas.drop('Año y Mes', axis=1)                                                               

colum_Servicemprdiv2 = [col for col in df_Servicempresas.columns if col not in ['PBI', 'Periodo']]

for col in colum_Servicemprdiv2:
    df_Servicempresas[col] = df_Servicempresas[col] * df_Servicempresas['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Servicempresas = df_Servicempresas.drop('PBI', axis=1) 

print(df_Servicempresas)


##EXCEL SERVICIOS GUBERNAMENTALES

df_Servicgubern = pd.read_excel(glob.glob('13*.xls')[0], sheet_name=0, usecols='A:B', skiprows=7)

df_Servicgubern.columns= ['Periodo','Servicios Gubernamentales']

df_Servicgubern['Periodo'] = df_Servicgubern['Periodo'].astype(int).astype(str)

nuevo_registro = {'Periodo': "XXXXXX", 'Servicios Gubernamentales': 0.0429255567059648}

df_Servicgubern.loc[len(df_Servicgubern)] = nuevo_registro

ult_registro_pondera = df_Servicgubern.iloc[-1]                                                                                   #Extraemos últimA fila, que posee los ponderadores 2007

excluye_column = 'Periodo'                                                                                                   #Excluímos la columna periodo

for col in df_Servicgubern.columns:
    if col != excluye_column:
        df_Servicgubern[col] = df_Servicgubern[col] * ult_registro_pondera[col]                                                          #Multiplicamos todos los registros por la última fila, excepto columna periodo                                 

df_Servicgubern.drop(index=df_Servicgubern.index[-1], inplace=True) 


df_Servicgubern = pd.merge(df_Servicgubern, df_VBP[['Año y Mes','Índice Global']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos la columna índice VBP PBI del excel correspondiente

df_Servicgubern = df_Servicgubern.drop('Año y Mes', axis=1)

colum_Servicguberndiv = [col for col in df_Servicgubern.columns if col not in ['Índice Global', 'Periodo']]

for col in colum_Servicguberndiv:
    df_Servicgubern[col] = (df_Servicgubern[col] / df_Servicgubern['Índice Global'])                                          #Dividimos todos los registros por el índice PBI para obtener participación del PBI en el periodo

df_Servicgubern.drop(columns=['Índice Global'], inplace=True)                                                         #Luego eliminamos columna índice VBP PBI

df_Servicgubern = pd.merge(df_Servicgubern, df_VAB[['Año y Mes','PBI']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')                                           #Traemos el VAB del excel correspondiente

df_Servicgubern = df_Servicgubern.drop('Año y Mes', axis=1)                                                               

colum_Servicguberndiv2 = [col for col in df_Servicgubern.columns if col not in ['PBI', 'Periodo']]

for col in colum_Servicguberndiv2:
    df_Servicgubern[col] = df_Servicgubern[col] * df_Servicgubern['PBI']                                                      #Multiplicamos a todas las columnas por el VAB

df_Servicgubern = df_Servicgubern.drop('PBI', axis=1) 

print(df_Servicgubern)


##EXCEL OTROS SERVICIOS
df_Otrosservicios = pd.merge(df_Servicgubern, df_Transp[['Transporte y Almacenamiento', 'Periodo']], 
                             left_on='Periodo', right_on='Periodo', how='left')

df_Otrosservicios = pd.merge(df_Otrosservicios, df_Alojrest[['Total_Alojrest', 'Periodo']], 
                             left_on='Periodo', right_on='Periodo', how='left')

df_Otrosservicios = pd.merge(df_Otrosservicios, df_Telecom[['Total_Telecomyotro', 'Periodo']], 
                             left_on='Periodo', right_on='Periodo', how='left')

df_Otrosservicios = pd.merge(df_Otrosservicios, df_Financi[['Servicios Financieros', 'Periodo']], 
                             left_on='Periodo', right_on='Periodo', how='left')

df_Otrosservicios = pd.merge(df_Otrosservicios, df_Servicempresas[['Servicios prestados a empresas', 'Periodo']], 
                             left_on='Periodo', right_on='Periodo', how='left')

df_Otrosservicios = pd.merge(df_Otrosservicios, df_VAB[['Otros servicios','Año y Mes']], 
                   left_on='Periodo', right_on='Año y Mes', how='left')   

df_Otrosservicios = df_Otrosservicios.drop('Año y Mes', axis=1) 

df_Otrosservicios = df_Otrosservicios.rename(columns={'Otros servicios': 'Total servicios'})

df_Otrosservicios['Otros servicios'] = df_Otrosservicios['Total servicios'] - df_Otrosservicios['Transporte y Almacenamiento']- df_Otrosservicios['Total_Alojrest']- df_Otrosservicios['Total_Telecomyotro']- df_Otrosservicios['Servicios Financieros']- df_Otrosservicios['Servicios prestados a empresas']- df_Otrosservicios['Servicios Gubernamentales']

df_Otrosservicios= df_Otrosservicios[['Periodo','Otros servicios']]

#-----Dinamizando columnas (producto, valor) y anexando sectores:

correla_sectores=pd.read_excel('G:/Mi unidad/CONSULTORA/Proyecto seguimiento economia/PBI Sectorial Peru/Script y correlacionadores/Correlacionador.xlsx', sheet_name='Correla', keep_default_na=False, usecols='A:G')                             #Importamos correlcionador CIIU4

print(correla_sectores)

print(df_agro)
df_agro = df_agro.drop('Agricultura, ganadería, caza y silvicultura', axis=1)
df_agro = df_agro.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_pesca)
df_pesca = df_pesca.drop('Pesca y acuicultura', axis=1)
df_pesca = df_pesca.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_minehidro)
df_minehidro = df_minehidro.drop('Extraccion de petróleo, gas, minerales y servicios conexos', axis=1)
df_minehidro = df_minehidro.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_manufac)
df_manufac = df_manufac.drop('Manufactura', axis=1)
df_manufac = df_manufac.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_electragua)
df_electragua = df_electragua.drop('Electricidad, gas, suministro de agua, alcantarillado y gestión de desechos y saneamiento', axis=1)
df_electragua = df_electragua.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_construcc)
df_construcc = df_construcc.drop('Construcción', axis=1)
df_construcc = df_construcc.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_comercio)
df_comercio = df_comercio.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_impuestos)
df_impuestos = df_impuestos.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Transp)
df_Transp = df_Transp.drop('Transporte y Almacenamiento', axis=1)
df_Transp = df_Transp.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Alojrest)
df_Alojrest = df_Alojrest.drop('Total_Alojrest', axis=1)
df_Alojrest = df_Alojrest.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Telecom)
df_Telecom = df_Telecom.drop('Total_Telecomyotro', axis=1)
df_Telecom = df_Telecom.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Financi)
df_Financi = df_Financi.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Servicempresas)
df_Servicempresas = df_Servicempresas.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Servicgubern)
df_Servicgubern = df_Servicgubern.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

print(df_Otrosservicios)
df_Otrosservicios = df_Otrosservicios.melt(id_vars='Periodo', var_name='Producto', value_name='Valor')

df_PBI_sectores = pd.concat([df_agro, df_pesca, df_minehidro,df_manufac,df_electragua,df_construcc,df_comercio,df_impuestos,df_Transp,df_Alojrest,df_Telecom,df_Financi,df_Servicempresas,df_Servicgubern,df_Otrosservicios], ignore_index=True)

print(df_PBI_sectores)

df_PBI_sectores = pd.merge(df_PBI_sectores, correla_sectores, 
                   left_on='Producto', right_on='Producto', how='left')

df_PBI_sectores.to_excel('BD_PBI_Sectorial.xlsx', sheet_name='Base',index=False)



