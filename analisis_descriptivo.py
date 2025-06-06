import matplotlib
# Usar el backend "Agg" para que no abra ventanas interactivas
matplotlib.use('Agg')

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
import os
from pathlib import Path
from typing import Dict, Any, Tuple
import logging


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('analisis_financiero.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuración de visualización
try:
    plt.style.use('seaborn-v0_8')
except OSError:
    plt.style.use('seaborn')
sns.set_palette("husl")
warnings.filterwarnings('ignore', category=FutureWarning)


class AnalisisFinanciero:
    """
    Clase principal para análisis financiero de datos de ventas.
    Incluye validaciones, análisis descriptivo, análisis avanzado y generación de informe Excel.
    """

    def __init__(self, ruta_archivo: str = 'Dataset.xlsx'):
        """
        Inicializa el análisis financiero.

        Args:
            ruta_archivo (str): Ruta al archivo Excel del dataset
        """
        self.ruta_archivo = Path(ruta_archivo)
        self.df: pd.DataFrame = pd.DataFrame()
        self.df_original: pd.DataFrame = pd.DataFrame()
        # Columnas mínimas requeridas:
        self.columnas_requeridas = ['Sales', 'Profit', 'Segment', 'Country', 'Date']

    def cargar_datos(self) -> bool:
        """
        Carga y valida el dataset con manejo robusto de errores.

        Returns:
            bool: True si la carga fue exitosa, False en caso contrario
        """
        try:
            logger.info(f"[OK] Intentando cargar archivo: {self.ruta_archivo}")

            if not self.ruta_archivo.exists():
                logger.error(f"Archivo no encontrado: {self.ruta_archivo}")
                return False

            # Cargar el archivo
            self.df = pd.read_excel(self.ruta_archivo)
            self.df_original = self.df.copy()

            # Limpiar nombres de columnas (eliminar espacios en blanco)
            self.df.columns = self.df.columns.str.strip()

            # Validar columnas requeridas
            if not self._validar_columnas():
                return False

            # Convertir fechas
            self._procesar_fechas()

            logger.info(f"[OK] Dataset cargado: {self.df.shape[0]} filas, {self.df.shape[1]} columnas")
            return True

        except Exception as e:
            logger.error(f"Error al cargar datos: {e}", exc_info=True)
            return False

    def _validar_columnas(self) -> bool:
        """
        Valida que existan las columnas requeridas.
        """
        columnas_faltantes = [col for col in self.columnas_requeridas if col not in self.df.columns]
        if columnas_faltantes:
            logger.error(f"Columnas faltantes: {columnas_faltantes}")
            return False
        return True

    def _procesar_fechas(self):
        """
        Procesa y valida la columna 'Date'.
        """
        try:
            if 'Date' in self.df.columns:
                # Si algunos valores no se pueden convertir, se convierten en NaT
                self.df['Date'] = pd.to_datetime(self.df['Date'], errors='coerce')
                n_fechas_invalidas = self.df['Date'].isna().sum()
                if n_fechas_invalidas > 0:
                    logger.warning(f"{n_fechas_invalidas} filas con 'Date' inválido se convirtieron a NaT")
                else:
                    logger.info("[OK] Fechas procesadas correctamente")
        except Exception as e:
            logger.warning(f"Advertencia al procesar fechas: {e}")

    def resumen_estructura(self) -> pd.DataFrame:
        """
        Genera un DataFrame con la estructura del dataset:
          - Nombre de columna
          - Tipo de dato (dtype)
          - Cantidad de valores nulos
        """
        if self.df.empty:
            logger.error("Datos no cargados. Ejecute cargar_datos() primero.")
            return pd.DataFrame()

        df_estructura = pd.DataFrame({
            'Columna': self.df.columns,
            'TipoDato': [str(dtype) for dtype in self.df.dtypes],
            'ValoresNulos': [int(self.df[col].isna().sum()) for col in self.df.columns]
        })

        logger.info("[OK] Resumen de estructura generado")
        return df_estructura

    def analisis_descriptivo_numerico(self) -> pd.DataFrame:
        """
        Realiza un análisis descriptivo básico de las columnas numéricas:
          - count, mean, std, min, 25%, 50%, 75%, max
        """
        if self.df.empty:
            logger.error("Datos no cargados. Ejecute cargar_datos() primero.")
            return pd.DataFrame()

        numeric_df = self.df.select_dtypes(include=[np.number])
        if numeric_df.empty:
            logger.warning("No hay columnas numéricas para describir.")
            return pd.DataFrame()

        df_descriptivo = numeric_df.describe().round(4).transpose().reset_index()
        df_descriptivo = df_descriptivo.rename(columns={'index': 'Columna'})

        logger.info("[OK] Análisis descriptivo numérico generado")
        return df_descriptivo

    def resumen(self) -> Dict[str, Any]:
        """
        Genera un resumen ejecutivo del dataset con métricas clave.

        Returns:
            Dict: Resumen con métricas clave
        """
        if self.df.empty:
            logger.error("Datos no cargados. Ejecute cargar_datos() primero.")
            return {}

        total_ventas = self.df['Sales'].sum()
        total_profit = self.df['Profit'].sum()
        margen_global = (total_profit / total_ventas * 100) if total_ventas > 0 else 0.0

        # Asegurarse de que no haya fechas faltantes para calcular el periodo
        fechas_validas = self.df['Date'].dropna()
        periodo = "N/A"
        if not fechas_validas.empty:
            periodo = f"{fechas_validas.min().strftime('%Y-%m')} a {fechas_validas.max().strftime('%Y-%m')}"

        resumen = {
            'Total_Ventas': round(total_ventas, 2),
            'Total_Profit': round(total_profit, 2),
            'Margen_Global_%': round(margen_global, 2),
            'Num_Transacciones': int(len(self.df)),
            'Num_Paises': int(self.df['Country'].nunique()),
            'Num_Segmentos': int(self.df['Segment'].nunique()),
            'Periodo': periodo,
            'Ticket_Promedio': round(self.df['Sales'].mean(), 2),
            'Transacciones_Negativas': int((self.df['Profit'] < 0).sum())
        }

        logger.info("[OK] Resumen ejecutivo generado")
        return resumen

    def calcular_margen_bruto(self) -> pd.DataFrame:
        """
        Calcula el margen bruto mejorado con validaciones adicionales.

        Retorna un DataFrame con la columna 'Margen_Bruto' agregada (en porcentaje).
        """
        if self.df.empty:
            logger.error("Datos no cargados.")
            return pd.DataFrame()

        df_work = self.df.copy()

        if 'Margen_Bruto' in df_work.columns:
            logger.info("[OK] La columna 'Margen_Bruto' ya existe. Se sobrescribirá con el nuevo cálculo.")

        # Cálculo con manejo de casos extremos
        condiciones = [
            df_work['Sales'] > 0,
            df_work['Sales'] == 0,
            df_work['Sales'] < 0
        ]
        opciones = [
            (df_work['Profit'] / df_work['Sales']) * 100,  # Caso normal
            0.0,                                           # Ventas cero → margen 0%
            np.nan                                        # Ventas negativas → NaN
        ]

        df_work['Margen_Bruto'] = np.select(condiciones, opciones, default=np.nan)

        # Estadísticas del margen
        margen_stats = df_work['Margen_Bruto'].describe()
        promedio_margen = margen_stats.get('mean', 0.0) if 'mean' in margen_stats else 0.0
        logger.info(f"[OK] Margen bruto calculado. Promedio: {promedio_margen:.2f}%")

        self.df = df_work
        return df_work

    def analisis_por_segmento(self) -> pd.DataFrame:
        """
        Análisis avanzado por segmento con métricas esenciales:
          - Segment
          - Total_Ventas
          - Beneficio_Total
          - Margen_Promedio

        Ordenado por Beneficio_Total descendente.
        """
        if self.df.empty:
            logger.error("Datos no cargados. Ejecute cargar_datos() primero.")
            return pd.DataFrame()

        # Asegurar que exista Margen_Bruto
        if 'Margen_Bruto' not in self.df.columns:
            logger.info("[OK] No existe 'Margen_Bruto'. Calculando...")
            self.calcular_margen_bruto()

        resumen = (
            self.df
            .groupby('Segment')
            .agg({
                'Sales': 'sum',
                'Profit': 'sum',
                'Margen_Bruto': 'mean'
            })
            .rename(columns={
                'Sales': 'Total_Ventas',
                'Profit': 'Beneficio_Total',
                'Margen_Bruto': 'Margen_Promedio'
            })
            .reset_index()
        )

        resumen = resumen.sort_values('Beneficio_Total', ascending=False).reset_index(drop=True)

        logger.info("[OK] Análisis por segmento completado")
        return resumen

    def calcular_metricas_pais(self, pais: str) -> Dict[str, Any]:
        """
        Calcula métricas avanzadas para un país específico.

        Args:
            pais (str): Nombre del país a analizar

        Returns:
            Dict: Métricas completas del país:
              - Ventas_Totales
              - Crecimiento_Intermensual_%
              - Margen_Promedio_%
        """
        if self.df.empty:
            logger.error("Datos no cargados.")
            return {}

        # Validar que el país exista en self.df
        paises_disponibles = self.df['Country'].unique()
        if pais not in paises_disponibles:
            logger.warning(f"País '{pais}' no encontrado. Disponibles: {', '.join(paises_disponibles)}")
            return {
                'Ventas_Totales': 0.0,
                'Crecimiento_Intermensual_%': 0.0,
                'Margen_Promedio_%': np.nan
            }

        # Asegurar Margen_Bruto
        if 'Margen_Bruto' not in self.df.columns:
            logger.info("[OK] No existe 'Margen_Bruto'. Calculando...")
            self.calcular_margen_bruto()

        # Filtrar por país
        df_pais = self.df[self.df['Country'] == pais].copy()
        if df_pais.empty:
            return {
                'Ventas_Totales': 0.0,
                'Crecimiento_Intermensual_%': 0.0,
                'Margen_Promedio_%': np.nan
            }

        # 1) Ventas totales
        ventas_totales = df_pais['Sales'].sum()
        # 2) Margen promedio
        margen_promedio = df_pais['Margen_Bruto'].mean() if 'Margen_Bruto' in df_pais.columns else np.nan
        # 3) Crecimiento intermensual promedio
        crecimiento_promedio = self._calcular_crecimiento_intermensual(df_pais)

        metricas = {
            'Ventas_Totales': round(ventas_totales, 2),
            'Crecimiento_Intermensual_%': round(crecimiento_promedio, 2),
            'Margen_Promedio_%': round(margen_promedio, 2) if not np.isnan(margen_promedio) else np.nan
        }

        logger.info(f"[OK] Métricas calculadas para {pais}")
        return metricas

    def _calcular_crecimiento_intermensual(self, df_pais: pd.DataFrame) -> float:
        """
        Calcula el crecimiento intermensual del país dado.

        Usa frecuencia 'M' (fin de mes) para agrupar ventas por mes.
        """
        try:
            df_fechas = df_pais.dropna(subset=['Date']).set_index('Date')
            if df_fechas.empty:
                return 0.0

            ventas_mensuales = (
                df_fechas['Sales']
                .resample('M')
                .sum()
                .sort_index()
            )
            if len(ventas_mensuales) < 2:
                return 0.0

            crecimiento = ventas_mensuales.pct_change().dropna()
            return float(crecimiento.mean() * 100) if len(crecimiento) > 0 else 0.0

        except Exception as e:
            logger.warning(f"Error calculando crecimiento intermensual: {e}")
            return 0.0

    def generar_heatmap_correlacion(self, ruta_salida: str = "heatmap_correlacion.png") -> None:
        """
        Genera un heatmap de correlación entre todas las variables numéricas y lo guarda como PNG.

        Args:
            ruta_salida (str): Ruta donde se guardará el archivo PNG del heatmap.
        """
        if self.df.empty:
            logger.error("Datos no cargados. No se puede generar el heatmap.")
            return

        numeric_df = self.df.select_dtypes(include=[np.number])
        if numeric_df.shape[1] < 2:
            logger.warning("No hay suficientes variables numéricas para generar el heatmap.")
            return

        corr = numeric_df.corr()

        plt.figure(figsize=(8, 6))
        sns.heatmap(
            corr,
            annot=True,
            fmt=".2f",
            cmap="RdYlBu_r",
            vmin=-1, vmax=1,
            linewidths=0.5,
            cbar_kws={"label": "Coeficiente de correlación"}
        )
        plt.xticks(rotation=45, ha="right")
        plt.yticks(rotation=0)
        plt.title("Heatmap de correlaciones (variables numéricas)", pad=12)
        plt.tight_layout()
        plt.savefig(ruta_salida, dpi=150, bbox_inches="tight")
        plt.close()
        logger.info(f"[OK] Heatmap de correlación guardado en '{ruta_salida}'")

    def exportar_informe_profesional(self, ruta: str = "Informe_Analisis_Financiero.xlsx") -> None:
        """
        Genera un archivo Excel profesional con las siguientes hojas (sin portada):
          1) Estructura
          2) Descriptivo Numérico
          3) Margen Bruto
          4) Análisis Segmento
          5) Métricas por País
          6) Mapa de Calor (inserta el PNG del heatmap)

        Args:
            ruta (str): Nombre/ruta del archivo Excel de salida.
        """
        if self.df.empty:
            logger.error("Datos no cargados.")
            return

        # Asegurar Margen_Bruto
        if 'Margen_Bruto' not in self.df.columns:
            self.calcular_margen_bruto()

        # 1) Estructura del dataset
        df_estructura = self.resumen_estructura()

        # 2) Análisis descriptivo de columnas numéricas
        df_descriptivo = self.analisis_descriptivo_numerico()

        # 3) Margen Bruto (solo Sales, Profit, Margen_Bruto)
        df_margen = self.df[['Sales', 'Profit', 'Margen_Bruto']].copy()

        # 4) Análisis por Segmento (solo tres columnas)
        df_segmentos = self.analisis_por_segmento().copy()

        # 5) Métricas por País
        metricas_paises = {}
        for pais in sorted(self.df['Country'].unique()):
            metricas_paises[pais] = self.calcular_metricas_pais(pais)
        df_metricas_paises = pd.DataFrame(metricas_paises).T
        df_metricas_paises.index.name = "Country"
        df_metricas_paises.reset_index(inplace=True)

        # 6) Matriz de Correlación Numérica
        numeric_df = self.df.select_dtypes(include=[np.number])
        corr_matrix = numeric_df.corr().round(4)

        # 7) Generar mapa de calor
        ruta_heatmap = "heatmap_correlacion.png"
        self.generar_heatmap_correlacion(ruta_heatmap)

        # Eliminar informe previo si existe
        if os.path.exists(ruta):
            try:
                os.remove(ruta)
            except PermissionError:
                logger.error(f"No se pudo eliminar '{ruta}'. Cierra el archivo y vuelve a intentar.")
                return

        with pd.ExcelWriter(ruta, engine='xlsxwriter') as writer:
            workbook = writer.book

            formato_encabezado = workbook.add_format({
                'bold': True,
                'bg_color': '#DCE6F1',
                'border': 1,
                'align': 'center'
            })
            formato_num = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            formato_entero = workbook.add_format({'num_format': '0', 'border': 1})
            formato_porcentaje = workbook.add_format({'num_format': '0.00%', 'border': 1})
            formato_texto = workbook.add_format({'border': 1})

            # ==== 7.1 HOJA ESTRUCTURA ====
            nombre_hoja_estructura = "Estructura"
            df_estructura.to_excel(writer, sheet_name=nombre_hoja_estructura, index=False, startrow=1)
            hoja_estructura = writer.sheets[nombre_hoja_estructura]

            for col_num, value in enumerate(df_estructura.columns.values):
                hoja_estructura.write(1, col_num, value, formato_encabezado)
                ancho = max(len(str(value)) + 2, 15)
                hoja_estructura.set_column(col_num, col_num, ancho)

                for fila in range(2, 2 + len(df_estructura)):
                    cell_val = df_estructura.at[fila - 2, value]
                    if value == "ValoresNulos":
                        hoja_estructura.write_number(fila, col_num, int(cell_val), formato_entero)
                    else:
                        hoja_estructura.write(fila, col_num, cell_val, formato_texto)

            hoja_estructura.autofilter(1, 0, 1 + len(df_estructura), len(df_estructura.columns) - 1)

            # ==== 7.2 HOJA DESCRIPTIVO NUMÉRICO ====
            nombre_hoja_desc_num = "Descriptivo Numérico"
            df_descriptivo.to_excel(writer, sheet_name=nombre_hoja_desc_num, index=False, startrow=1)
            hoja_desc_num = writer.sheets[nombre_hoja_desc_num]

            for col_num, value in enumerate(df_descriptivo.columns.values):
                hoja_desc_num.write(1, col_num, value, formato_encabezado)
                ancho = max(len(str(value)) + 2, 15)
                hoja_desc_num.set_column(col_num, col_num, ancho)

                for fila in range(2, 2 + len(df_descriptivo)):
                    cell_val = df_descriptivo.at[fila - 2, value]
                    if isinstance(cell_val, (int, np.integer, float)):
                        hoja_desc_num.write_number(fila, col_num, cell_val, formato_num)
                    else:
                        hoja_desc_num.write(fila, col_num, cell_val, formato_texto)

            hoja_desc_num.autofilter(1, 0, 1 + len(df_descriptivo), len(df_descriptivo.columns) - 1)

            # ==== 7.3 HOJA MARGEN BRUTO ====
            nombre_hoja_margen = "Margen Bruto"
            df_margen.to_excel(writer, sheet_name=nombre_hoja_margen, index=False, startrow=1)
            hoja_margen = writer.sheets[nombre_hoja_margen]

            for col_num, value in enumerate(df_margen.columns.values):
                hoja_margen.write(1, col_num, value, formato_encabezado)
                ancho = max(len(str(value)) + 2, 15)
                hoja_margen.set_column(col_num, col_num, ancho)

                for fila in range(2, 2 + len(df_margen)):
                    cell_val = df_margen.at[fila - 2, value]
                    if value in ["Sales", "Profit"]:
                        hoja_margen.write_number(fila, col_num, cell_val, formato_num)
                    elif value == "Margen_Bruto":
                        hoja_margen.write_number(fila, col_num, cell_val / 100, formato_porcentaje)
                    else:
                        hoja_margen.write(fila, col_num, cell_val, formato_entero)

            hoja_margen.autofilter(1, 0, 1 + len(df_margen), len(df_margen.columns) - 1)

            # ==== 7.4 HOJA ANÁLISIS POR SEGMENTO ====
            nombre_hoja_seg = "Análisis Segmento"
            df_segmentos.to_excel(writer, sheet_name=nombre_hoja_seg, index=False, startrow=1)
            hoja_seg = writer.sheets[nombre_hoja_seg]

            for col_num, value in enumerate(df_segmentos.columns.values):
                hoja_seg.write(1, col_num, value, formato_encabezado)
                ancho = max(len(str(value)) + 2, 15)
                hoja_seg.set_column(col_num, col_num, ancho)
                for fila in range(2, 2 + len(df_segmentos)):
                    cell_val = df_segmentos.at[fila - 2, value]
                    if value in ["Total_Ventas", "Beneficio_Total"]:
                        hoja_seg.write_number(fila, col_num, cell_val, formato_num)
                    elif value == "Margen_Promedio":
                        hoja_seg.write_number(fila, col_num, cell_val / 100, formato_porcentaje)
                    else:  # Segment
                        hoja_seg.write(fila, col_num, cell_val)

            hoja_seg.autofilter(1, 0, 1 + len(df_segmentos), len(df_segmentos.columns) - 1)

            # ==== 7.5 HOJA MÉTRICAS POR PAÍS ====
            nombre_hoja_paises = "Métricas por País"
            df_metricas_paises.to_excel(writer, sheet_name=nombre_hoja_paises, index=False, startrow=1)
            hoja_paises = writer.sheets[nombre_hoja_paises]

            for col_num, value in enumerate(df_metricas_paises.columns.values):
                hoja_paises.write(1, col_num, value, formato_encabezado)
                ancho = max(len(str(value)) + 2, 15)
                hoja_paises.set_column(col_num, col_num, ancho)
                for fila in range(2, 2 + len(df_metricas_paises)):
                    cell_val = df_metricas_paises.at[fila - 2, value]
                    if value == "Ventas_Totales":
                        hoja_paises.write_number(fila, col_num, cell_val, formato_num)
                    elif value == "Crecimiento_Intermensual_%":
                        hoja_paises.write_number(fila, col_num, cell_val / 100, formato_porcentaje)
                    elif value == "Margen_Promedio_%":
                        hoja_paises.write_number(fila, col_num, cell_val / 100, formato_porcentaje)
                    else:  # Country
                        hoja_paises.write(fila, col_num, cell_val)

            hoja_paises.autofilter(1, 0, 1 + len(df_metricas_paises), len(df_metricas_paises.columns) - 1)

            # ==== 7.6 HOJA MAPA DE CALOR ====
            hoja_calor = workbook.add_worksheet("Mapa de Calor")
            writer.sheets["Mapa de Calor"] = hoja_calor
            hoja_calor.insert_image("A2", ruta_heatmap, {'x_scale': 0.7, 'y_scale': 0.7})

        logger.info(f"[OK] Informe profesional guardado en '{ruta}'")

    def generar_dashboard(self, pais_ejemplo: str = None) -> None:
        """
        Genera un dashboard completo con múltiples visualizaciones.

        Args:
            pais_ejemplo (str): País para análisis detallado (actualmente no usado)
        """
        if self.df.empty:
            logger.error("Datos no cargados.")
            return

        fig = plt.figure(figsize=(20, 16))
        gs = fig.add_gridspec(4, 3, hspace=0.3, wspace=0.3)

        # 1. Distribución de ventas
        ax1 = fig.add_subplot(gs[0, 0])
        self.df['Sales'].hist(bins=30, alpha=0.7, ax=ax1)
        ax1.set_title('Distribución de Ventas', fontweight='bold')
        ax1.set_xlabel('Ventas')
        ax1.set_ylabel('Frecuencia')

        # 2. Boxplot de profit por segmento
        ax2 = fig.add_subplot(gs[0, 1])
        sns.boxplot(data=self.df, x='Segment', y='Profit', ax=ax2)
        ax2.set_title('Distribución de Profit por Segmento', fontweight='bold')
        ax2.tick_params(axis='x', rotation=45)

        # 3. Ventas por mes (si existe la columna 'Month Name')
        ax3 = fig.add_subplot(gs[0, 2])
        if 'Month Name' in self.df.columns:
            monthly_sales = self.df.groupby('Month Name')['Sales'].sum()
            monthly_sales.plot(kind='bar', ax=ax3)
            ax3.set_title('Ventas por Mes', fontweight='bold')
            ax3.tick_params(axis='x', rotation=45)

        # 4. Top países por ventas
        ax4 = fig.add_subplot(gs[1, :])
        top_paises = self.df.groupby('Country')['Sales'].sum().sort_values(ascending=False).head(10)
        top_paises.plot(kind='bar', ax=ax4)
        ax4.set_title('Top 10 Países por Ventas Totales', fontweight='bold')
        ax4.tick_params(axis='x', rotation=45)

        # 5. Análisis de segmentos (ventas y margen) — opcional en dashboard
        ax5 = fig.add_subplot(gs[2, :])
        if 'Margen_Bruto' in self.df.columns:
            segment_analysis = (
                self.df.groupby('Segment')
                       .agg({'Sales': 'sum', 'Profit': 'sum', 'Margen_Bruto': 'mean'})
            )
            ax5_twin = ax5.twinx()
            segment_analysis['Sales'].plot(kind='bar', ax=ax5, alpha=0.7, color='skyblue')
            segment_analysis['Margen_Bruto'].plot(kind='line', ax=ax5_twin, color='red', marker='o')

            ax5.set_title('Ventas y Margen por Segmento', fontweight='bold')
            ax5.set_ylabel('Ventas', color='skyblue')
            ax5_twin.set_ylabel('Margen Bruto %', color='red')
            ax5.tick_params(axis='x', rotation=45)

        # 6. Evolución temporal de ventas y profit
        ax6 = fig.add_subplot(gs[3, :])
        ventas_periodo = (
            self.df.dropna(subset=['Date'])
            .set_index('Date')
            .resample('M')
            .agg({'Sales': 'sum', 'Profit': 'sum'})
        )
        if not ventas_periodo.empty:
            ventas_periodo.plot(ax=ax6)
            ax6.set_title('Evolución Temporal de Ventas y Profit', fontweight='bold')
            ax6.legend()

        plt.suptitle('Dashboard Financiero - Análisis Completo', fontsize=16, fontweight='bold')
        logger.info("[OK] Dashboard generado exitosamente")

    def exportar_resultados(self, directorio: str = "resultados_analisis") -> None:
        """
        Exporta todos los resultados a archivos organizados.

        Args:
            directorio (str): Directorio donde guardar los resultados
        """
        if self.df.empty:
            logger.error("Datos no cargados.")
            return

        Path(directorio).mkdir(exist_ok=True)

        try:
            # 1. Estructura del dataset
            df_estructura = self.resumen_estructura()
            df_estructura.to_csv(f"{directorio}/estructura_dataset.csv", index=False)

            # 2. Análisis descriptivo numérico
            df_descriptivo = self.analisis_descriptivo_numerico()
            df_descriptivo.to_csv(f"{directorio}/descriptivo_numerico.csv", index=False)

            # 3. Resumen ejecutivo
            resumen = self.resumen()
            pd.DataFrame([resumen]).to_csv(f"{directorio}/resumen_ejecutivo.csv", index=False)

            # 4. Análisis por segmento
            if 'Margen_Bruto' not in self.df.columns:
                self.calcular_margen_bruto()
            analisis_segmento = self.analisis_por_segmento()
            analisis_segmento.to_csv(f"{directorio}/analisis_segmentos.csv", index=False)

            # 5. Dataset completo procesado
            self.df.to_csv(f"{directorio}/dataset_procesado.csv", index=False)

            # 6. Matriz de correlación + heatmap (solo heatmap)
            numeric_df = self.df.select_dtypes(include=[np.number])
            corr_matrix = numeric_df.corr().round(4)
            corr_matrix.to_csv(f"{directorio}/matriz_correlacion.csv")
            # El heatmap ya se guardó como "heatmap_correlacion.png"
            logger.info(f"[OK] Heatmap de correlación guardado como 'heatmap_correlacion.png'")

            # 7. Métricas por país
            metricas_paises = {}
            for pais in sorted(self.df['Country'].unique()):
                metricas_paises[pais] = self.calcular_metricas_pais(pais)
            pd.DataFrame(metricas_paises).T.to_csv(f"{directorio}/metricas_por_pais.csv")

            logger.info(f"[OK] Resultados exportados a: {directorio}")

        except Exception as e:
            logger.error(f"Error exportando resultados: {e}", exc_info=True)


def main():
    """
    Función principal para ejecutar el análisis completo.
    """
    print("INICIANDO ANÁLISIS FINANCIERO AVANZADO")
    print("=" * 50)

    # Inicializar análisis
    analisis = AnalisisFinanciero('Dataset.xlsx')

    # 1) Cargar datos
    if not analisis.cargar_datos():
        print("Error al cargar datos. Terminando ejecución.")
        return

    # 2) Mostrar resumen de estructura y descriptivo numérico en consola 
    print("\nESTRUCTURA DEL DATASET")
    print("-" * 30)
    df_estructura = analisis.resumen_estructura()
    print(df_estructura.to_string(index=False))

    print("\nANÁLISIS DESCRIPTIVO DE COLUMNAS NUMÉRICAS")
    print("-" * 30)
    df_descriptivo = analisis.analisis_descriptivo_numerico()
    if not df_descriptivo.empty:
        print(df_descriptivo.to_string(index=False))
    else:
        print("No hay columnas numéricas para describir.")

    # 3) Calcular margen bruto
    print("\nCALCULANDO MARGEN BRUTO...")
    analisis.calcular_margen_bruto()

    # 4) Análisis por segmento en consola
    print("\nANÁLISIS POR SEGMENTO")
    print("-" * 30)
    segmentos = analisis.analisis_por_segmento()
    if not segmentos.empty:
        print(segmentos.to_string(index=False))
    else:
        print("No se pudo generar el análisis por segmento.")

    # 5) Análisis de país (ejemplo con 'France') en consola
    print("\nANÁLISIS POR PAÍS (France)")
    print("-" * 30)
    metricas_francia = analisis.calcular_metricas_pais('France')
    for clave, valor in metricas_francia.items():
        if isinstance(valor, (int, np.integer)):
            print(f"{clave}: {valor:d}")
        elif isinstance(valor, float):
            if valor.is_integer():
                print(f"{clave}: {int(valor):d}")
            else:
                print(f"{clave}: {valor:,.2f}")
        else:
            print(f"{clave}: {valor}")

    # 6) Análisis de correlación (solo heatmap) en consola
    print("\nANÁLISIS DE CORRELACIÓN")
    print("-" * 30)
    analisis.generar_heatmap_correlacion("heatmap_correlacion.png")
    print("Heatmap de correlación guardado en 'heatmap_correlacion.png'.")

    # 7) Exportar informe profesional a Excel (sin portada)
    print("\nEXPORTANDO INFORME PROFESIONAL A EXCEL…")
    # Si existe, eliminar antes de escribir
    if os.path.exists("Informe_Analisis_Financiero.xlsx"):
        try:
            os.remove("Informe_Analisis_Financiero.xlsx")
        except PermissionError:
            print("Cierra 'Informe_Analisis_Financiero.xlsx' y vuelve a ejecutar.")
            return

    analisis.exportar_informe_profesional("Informe_Analisis_Financiero.xlsx")
    print("Archivo 'Informe_Analisis_Financiero.xlsx' generado con formato profesional.")
    print("\nANÁLISIS COMPLETADO EXITOSAMENTE")
    print("=" * 50)


if __name__ == "__main__":
    main()
