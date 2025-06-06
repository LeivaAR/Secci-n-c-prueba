# Análisis Financiero Avanzado
## Requisitos Previos

- Python 3.8 o superior  
- Las siguientes bibliotecas de Python (se instalan desde `requirements.txt`):
pandas>=1.3.0
numpy>=1.19.0
matplotlib>=3.4.0
seaborn>=0.11.0
openpyxl>=3.0.0
xlsxwriter>=3.0.0


---

## Pasos para ejecutar localmente

1. **Descomprimir**  
 Extrae el contenido del archivo comprimido (ZIP, tar.gz, etc.) en una carpeta local.

2. **Crear y activar entorno virtual** (opcional, pero recomendado)  
 - En Linux/macOS:
   ```bash
   python3 -m venv env
   source env/bin/activate
   ```
 - En Windows PowerShell:
   ```powershell
   python -m venv env
   .\env\Scripts\Activate.ps1
   ```
 - En Windows CMD:
   ```cmd
   python -m venv env
   env\Scripts\activate.bat
   ```

3. **Instalar dependencias**  
 Navega a la carpeta donde extrajiste todo y ejecuta:
 ```bash
 pip install -r requirements.txt

Instalación y Ejecución Local
Sigue estos pasos para preparar el proyecto en tu máquina:

Clonar o descargar el repositorio

Si ya tienes un ZIP: descomprímelo en la carpeta que prefieras.

Si prefieres clonar directamente abre una terminal o PowerShell y ejecuta:

## Instalación y Ejecución Local

Sigue estos pasos para preparar el proyecto en tu máquina:

1. **Clonar o descargar el repositorio**  
   - Si ya tienes un ZIP:  
     - Descomprímelo en la carpeta que prefieras.  
   - Si prefieres clonar directamente, abre una terminal o PowerShell y ejecuta:
     ```bash
     git clone https://github.com/LeivaAR/Secci-n-c-prueba.git
     cd Secci-n-c-prueba
     ```

2. **Crear y activar un entorno virtual** (opcional, pero recomendado)  
   Un entorno virtual te asegura que las dependencias no interfieran con otros proyectos de Python que tengas en la máquina.
   - **En Linux/macOS**:
     ```bash
     python3 -m venv env
     source env/bin/activate
     ```
   - **En Windows PowerShell**:
     ```powershell
     python -m venv env
     .\env\Scripts\Activate.ps1
     ```
   - **En Windows (CMD)**:
     ```cmd
     python -m venv env
     env\Scripts\activate.bat
     ```
   Después de esto, verás el prefijo `(env)` en tu terminal, lo que indica que estás usando el entorno virtual.

3. **Instalar dependencias**  
   Con el entorno (env) activo, navega a la carpeta donde quedó el proyecto (si no lo hiciste en el paso anterior) y ejecuta:
   ```bash
   pip install -r requirements.txt


4. Ejecutar los scripts principales

Para probar la carga de datos y el preprocesamiento, por ejemplo:

python src/cargar_datos.py

Para generar un análisis de portafolio completo:

python src/analisis_portafolio.py

Para generar gráficos de tendencias:

python src/graficos/plot_trends.py
 
Cada script suele imprimir en consola el estado de avance y, en la mayoría de casos, guardará resultados (tablas, gráficos) en carpetas de salida (por ejemplo, outputs/ o reportes/ si es que las creaste).

5. Desactivar el entorno cuando termines
Solo escribe:

deactivate