import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
from datetime import datetime

def consultar_nits(input_path, output_path):
    df = pd.read_excel(input_path, dtype={'NIT': str})

    options = uc.ChromeOptions()
    options.add_argument('--no-first-run')
    options.add_argument('--no-service-autorun')
    options.add_argument('--password-store=basic')

    driver = uc.Chrome(options=options)
    driver.get("https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces")

    print("🚨 Por favor, resuelve el CAPTCHA manualmente. El programa continuará después de 15 segundos")
    time.sleep(8)  # Espera de 15s. Cambia esto si es necesario.

    def safe_get_text(by, value):
        try:
            return driver.find_element(by, value).text
        except NoSuchElementException:
            return ''

    resultados = []
    fallidos = []

    for index, row in df.iterrows():
        nit = row['NIT'].split('.')[0].strip()
        print(f'----Consultando NIT: {nit}')

        try:
            time.sleep(5)

            input_nit = driver.find_element(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit')
            input_nit.clear()
            time.sleep(1)
            input_nit.send_keys(nit)

            driver.find_element(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar').click()
            time.sleep(5)

            razon_social = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:razonSocial')
            dv = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv')
            fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            estado = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado')
            primer_nombre = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre')
            segundo_nombre = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres')
            primer_apellido = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido')
            segundo_apellido = safe_get_text(By.ID, 'vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido')

            resultados.append({
                'NIT': nit or 'No aplica',
                'Razón Social': razon_social or 'No aplica',
                'DV': dv or 'No aplica',
                'Primer Nombre': primer_nombre or 'No aplica',
                'Segundo Nombre': segundo_nombre or 'No aplica',
                'Primer Apellido': primer_apellido or 'No aplica',
                'Segundo Apellido': segundo_apellido or 'No aplica',
                'Fecha Busqueda': fecha_actual or 'No aplica',
                'Estado RUT': estado or 'Error de consulta'
            })
            print(f'✅ Resultado obtenido para {nit}')

        except Exception as e:
            print(f'❌ Error con NIT {nit}: {e}')
            fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            error_data = {
                'NIT': nit,
                'Razón Social': 'NA',
                'DV': '0',
                'Primer Nombre': 'NA',
                'Segundo Nombre': 'NA',
                'Primer Apellido': 'NA',
                'Segundo Apellido': 'NA',
                'Fecha Busqueda': fecha_actual,
                'Estado RUT': 'No está registrado en la DIAN o error de consulta.'
            }
            resultados.append(error_data)
            fallidos.append(error_data)

        driver.get("https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces")
        time.sleep(2)

    driver.quit()

    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel(output_path, index=False)
    return output_path
