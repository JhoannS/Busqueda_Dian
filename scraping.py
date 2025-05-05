import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
from datetime import datetime
from tqdm import tqdm
import logging
import os

logging.basicConfig(filename='log_consulta.log', level=logging.INFO)

def safe_get_text(by, value, driver):
    try:
        return driver.find_element(by, value).text
    except NoSuchElementException:
        return ''

def esperar_capcha(driver):
    print("🧠 Esperando a que se muestre el campo de NIT (posiblemente después del CAPTCHA)...")
    WebDriverWait(driver, timeout=1000).until(
        EC.presence_of_element_located((By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit'))
    )
    print("✔️ Página lista para ingresar NIT.")

def esperar_validacion_captcha(driver):
    print("🧠 Verificando si el CAPTCHA ha sido validado correctamente...")
    while True:
        try:
            mensaje = driver.find_element(By.CLASS_NAME, 'ui-messages-error-summary').text
            if "captcha" in mensaje.lower():
                print("⚠️ CAPTCHA no resuelto aún. Esperando 5 segundos...")
                time.sleep(5)
            else:
                break
        except NoSuchElementException:
            break

def consultar_nits(file_path_input='consulta_nits.xlsx', file_path_output='resultados_dian.xlsx'):
    df = pd.read_excel(file_path_input, dtype={'NIT': str})
    resultados = []

    ya_hechos = set()
    if os.path.exists(file_path_output):
        df_existente = pd.read_excel(file_path_output, dtype={'NIT': str})
        ya_hechos = set(df_existente['NIT'].astype(str))
        resultados.extend(df_existente.to_dict('records'))

    options = uc.ChromeOptions()
    options.add_argument('--window-size=500,700')
    driver = uc.Chrome(options=options)

    for index, row in tqdm(df.iterrows(), total=len(df)):
        nit = row['NIT'].split('.')[0].strip()
        if nit in ya_hechos:
            continue

        for intento in range(3):
            try:
                print(f'🔄 Recargando página para consultar NIT: {nit} (Intento {intento+1})')
                driver.get("https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces")
                esperar_capcha(driver)

                print("⌛ Esperando 6 segundos para simular comportamiento humano...")
                time.sleep(6)

                input_nit = driver.find_element(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit')
                input_nit.clear()
                input_nit.send_keys(nit)

                print("🧠 Por favor, resuelve el CAPTCHA si aún no lo has hecho...")
                WebDriverWait(driver, timeout=1000).until(
                    EC.presence_of_element_located((By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar'))
                )

                driver.find_element(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar').click()
                esperar_validacion_captcha(driver)

                WebDriverWait(driver, timeout=20).until(
                    EC.presence_of_element_located((By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:razonSocial'))
                )

                razon_social = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:razonSocial', driver)
                dv = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv', driver)
                estado = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado', driver)
                primer_nombre = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre', driver)
                segundo_nombre = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres', driver)
                primer_apellido = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido', driver)
                segundo_apellido = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido', driver)
                fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                resultado = {
                    'NIT': nit or 'No aplica',
                    'Razón Social': razon_social or 'No aplica',
                    'DV': dv or 'No aplica',
                    'Primer Nombre': primer_nombre or 'No aplica',
                    'Segundo Nombre': segundo_nombre or 'No aplica',
                    'Primer Apellido': primer_apellido or 'No aplica',
                    'Segundo Apellido': segundo_apellido or 'No aplica',
                    'Fecha Busqueda': fecha_actual,
                    'Estado RUT': estado or 'No aplica'
                }

                resultados.append(resultado)
                logging.info(f"Consulta exitosa NIT: {nit} — {razon_social}, Estado: {estado}")
                print(f"✅ Datos capturados para NIT: {nit} — Razón Social: {razon_social}, Estado: {estado}")

                break  # salir del ciclo de reintentos si fue exitoso

            except Exception as e:
                logging.error(f'Error con NIT {nit} (Intento {intento+1}): {e}')
                print(f'❌ Error con NIT {nit} (Intento {intento+1}): {e}')
                time.sleep(3)

                if intento == 2:
                    resultados.append({
                        'NIT': nit,
                        'Razón Social': 'NA',
                        'DV': '0',
                        'Primer Nombre': 'NA',
                        'Segundo Nombre': 'NA',
                        'Primer Apellido': 'NA',
                        'Segundo Apellido': 'NA',
                        'Fecha Busqueda': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'Estado RUT': 'Error de consulta o no registrado'
                    })

        if index % 10 == 0:
            pd.DataFrame(resultados).to_excel(file_path_output, index=False)

    driver.quit()
    pd.DataFrame(resultados).to_excel(file_path_output, index=False)
    print(f"📁 Archivo guardado correctamente en: {file_path_output}")
    return file_path_output