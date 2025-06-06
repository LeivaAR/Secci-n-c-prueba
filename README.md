# Análisis Financiero Avanzado

## Requisitos Previos

- Python 3.8 o superior  
- Las siguientes bibliotecas de Python (se instalan desde `requirements.txt`):

```
pandas>=1.3.0
numpy>=1.19.0
matplotlib>=3.4.0
seaborn>=0.11.0
openpyxl>=3.0.0
xlsxwriter>=3.0.0
```

---

## Instalación y Ejecución Local

Sigue estos pasos para preparar el proyecto en tu máquina:

### 1. Clonar o descargar el repositorio

- **Si ya tienes un ZIP:**  
  Descomprímelo en la carpeta que prefieras.  

- **Si prefieres clonar directamente:**  
  Abre una terminal o PowerShell y ejecuta:
  ```bash
  git clone https://github.com/LeivaAR/Secci-n-c-prueba.git
  cd Secci-n-c-prueba
  ```

### 2. Crear y activar un entorno virtual (opcional, pero recomendado)

Un entorno virtual te asegura que las dependencias no interfieran con otros proyectos de Python que tengas en la máquina.

- **En Linux/macOS:**
  ```bash
  python3 -m venv env
  source env/bin/activate
  ```

- **En Windows PowerShell:**
  ```powershell
  python -m venv env
  .\env\Scripts\Activate.ps1
  ```

- **En Windows (CMD):**
  ```cmd
  python -m venv env
  env\Scripts\activate.bat
  ```

Después de esto, verás el prefijo `(env)` en tu terminal, lo que indica que estás usando el entorno virtual.

### 3. Instalar dependencias

Con el entorno (env) activo, navega a la carpeta donde quedó el proyecto (si no lo hiciste en el paso anterior) y ejecuta:

```bash
pip install -r requirements.txt
```

### 4. Ejecutar los scripts principales

Para probar la carga de datos y el preprocesamiento, por ejemplo:

```bash
python src/cargar_datos.py
```

Para generar un análisis de portafolio completo:

```bash
python .\analisis_descriptivo.py 
```

Para generar gráficos de tendencias:

```bash
python src/graficos/plot_trends.py
```

Cada script suele imprimir en consola el estado de avance y, en la mayoría de casos, guardará resultados (tablas, gráficos) en carpetas de salida (por ejemplo, `outputs/` o `reportes/` si es que las creaste).

### 5. Desactivar el entorno cuando termines

Solo escribe:

```bash
deactivate
```

---

## Estructura del Proyecto

```
Secci-n-c-prueba/
├── src/
│   ├── cargar_datos.py
│   ├── analisis_portafolio.py
│   └── graficos/
│       └── plot_trends.py
├── requirements.txt
├── outputs/
└── README.md
```

## Notas Adicionales

- Asegúrate de tener todos los archivos de datos necesarios en las carpetas correspondientes antes de ejecutar los scripts.
- Los resultados se guardarán en las carpetas de salida especificadas en cada script.
- Si encuentras algún error durante la instalación o ejecución, verifica que tengas la versión correcta de Python y que todas las dependencias se hayan instalado correctamente.