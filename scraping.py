import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
from datetime import datetime

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
            # No hay mensaje de error, continuar
            break

def consultar_nits(file_path_input='consulta_nits.xlsx', file_path_output='resultados_dian.xlsx'):
    df = pd.read_excel(file_path_input, dtype={'NIT': str})
    resultados = []

    options = uc.ChromeOptions()
    options.add_argument('--window-size=500,700')
    driver = uc.Chrome(options=options)

    for index, row in df.iterrows():
        nit = row['NIT'].split('.')[0].strip()
        print(f'🔄 Recargando página para consultar NIT: {nit}')
        driver.get("https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces")

        try:
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

            # Verificar y esperar hasta que el mensaje de captcha desaparezca
            esperar_validacion_captcha(driver)

            # Esperar los resultados
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

            resultados.append({
                'NIT': nit or 'No aplica',
                'Razón Social': razon_social or 'No aplica',
                'DV': dv or 'No aplica',
                'Primer Nombre': primer_nombre or 'No aplica',
                'Segundo Nombre': segundo_nombre or 'No aplica',
                'Primer Apellido': primer_apellido or 'No aplica',
                'Segundo Apellido': segundo_apellido or 'No aplica',
                'Fecha Busqueda': fecha_actual,
                'Estado RUT': estado or 'No aplica'
            })

            print(f"✅ Datos capturados para NIT: {nit} — Razón Social: {razon_social}, Estado: {estado}")

        except Exception as e:
            print(f'❌ Error con NIT {nit}: {e}')
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

    driver.quit()

    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel(file_path_output, index=False)
    print(f"📁 Archivo guardado correctamente en: {file_path_output}")
    return file_path_output
