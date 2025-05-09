# Consulta de NITs DIAN JoyPegasus

Este proyecto permite consultar información del **RUT** en la página oficial de la DIAN de Colombia, utilizando **Selenium** y **undetected-chromedriver** para evitar bloqueos por parte del sitio.

---

## Requisitos

- Python 3.8 o superior
- Google Chrome instalado
- pip (gestor de paquetes de Python)

---

## Instalación

1. **Clona este repositorio** (o copia los archivos en tu máquina):

- git clone https://github.com/JhoannS/Busqueda_Dian.git e ingresa con **cd consulta_dian_web**

2. **Crea un entorno virtual** 
python -m venv venv

3. **Activa el entorno virtual:**

- En GitBash: source venv/bin/activate 

- En PowerSheel: venv\Scripts\activate

3. **Instala dependencias**
- pip install -r requirements.txt (opcional)
- pip install pandas selenium openpyxl undetected-chromedriver flask setuptools (obligatorio)
