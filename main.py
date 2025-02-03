import PySimpleGUI as sg
import tabula
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import numpy as np
from fpdf import FPDF
import tempfile
import os

sg.theme('LightGrey1')

FONT_TITLE = ('Helvetica', 16, 'bold')
FONT_SUBTITLE = ('Helvetica', 12, 'bold')
FONT_BODY = ('Helvetica', 10)

def process_pdf(file_path):
    try:
        tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
        if not tables:
            return None, "No tables found in PDF"
        
        df = tables[0].copy()
        
        if len(df.columns) == 4:
            df.columns = ['Lectura', 'Hora', 'Voltaje', 'Ordenado_ascendente']
        elif len(df.columns) == 3:
            df.columns = ['Hora', 'Voltaje', 'Ordenado_ascendente']
        else:
            return None, "Unexpected table structure"
        
        df['Voltaje'] = pd.to_numeric(df['Voltaje'], errors='coerce')
        df = df.dropna(subset=['Voltaje'])
        df['Hora'] = df['Hora'].astype(str).str.replace(r'[^0-9:]', ':', regex=True)
        
        if df.empty:
            return None, "No valid voltage data found"
            
        return df, None
    except Exception as e:
        return None, str(e)

def calculate_statistics(df):
    real_voltage = 127.00
    stats = {}
    
    df['Error Absoluto'] = abs(df['Voltaje'] - real_voltage)
    df['Error Relativo (%)'] = (df['Error Absoluto'] / real_voltage) * 100
    
    stats['mean'] = df['Voltaje'].mean()
    stats['median'] = df['Voltaje'].median()
    stats['mode'] = df['Voltaje'].mode().tolist()
    stats['std_dev'] = df['Voltaje'].std()
    stats['variance'] = df['Voltaje'].var()
    stats['range'] = df['Voltaje'].max() - df['Voltaje'].min()
    stats['coefficient_variation'] = stats['std_dev'] / stats['mean'] if stats['mean'] != 0 else 0
    q1, q3 = df['Voltaje'].quantile([0.25, 0.75])
    stats['semi_interquartil'] = (q3 - q1) / 2
    stats['mad'] = (df['Voltaje'] - stats['mean']).abs().mean()
    
    return stats, df

def create_histogram(data):
    plt.figure(figsize=(8, 4))
    plt.hist(data, bins=12, edgecolor='black', alpha=0.7, color='#1f77b4')
    plt.title('Distribución de Lecturas de Voltaje')
    plt.xlabel('Voltaje (V)')
    plt.ylabel('Frecuencia')
    plt.axvline(127, color='red', linestyle='dashed', linewidth=1, label='Valor Nominal')
    plt.grid(True)
    plt.legend()
    
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150)
    plt.close()
    return buf.getvalue()

def create_report(stats, histogram_path, output_path):
    pass

def main_window():
    layout = [
        [sg.Push(), sg.Text('Analizador de Voltaje ISC8 Inc.', font=FONT_TITLE), sg.Push()],
        [sg.HorizontalSeparator()],
        [sg.Text('Seleccione el archivo PDF con las lecturas:', font=FONT_SUBTITLE, pad=(0, 20))],
        [
            sg.Input(key='-FILE-', size=(40, 1), font=FONT_BODY),
            sg.FileBrowse('Examinar', file_types=(("PDF Files", "*.pdf"),))
        ],
        [
            sg.Push(),
            sg.Button('Procesar', size=(10, 1), font=FONT_BODY),
            sg.Button('Salir', size=(10, 1), font=FONT_BODY),
            sg.Push()
        ]
    ]
    return sg.Window('Analizador de Voltaje', layout, element_justification='center')

def results_window(stats, df, histogram_data):
    data_table = df[['Hora', 'Voltaje', 'Error Absoluto', 'Error Relativo (%)']].values.tolist()
    headers_table = ['Hora', 'Voltaje (V)', 'Error Abs. (V)', 'Error Rel. (%)']
    
    stats_layout = [
        [sg.Text('Estadísticas Principales', font=FONT_SUBTITLE)],
        [sg.Column([
            [
                sg.Frame('Centralidad', [
                    [sg.Text('Media:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['mean']:.4f} V", font=('Courier New', 10))],
                    [sg.Text('Mediana:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['median']:.4f} V", font=('Courier New', 10))],
                    [sg.Text('Moda:', font=FONT_BODY), sg.Push(), sg.Text(', '.join(map(str, stats['mode'])) + ' V', font=('Courier New', 10))]
                ], border_width=0),
                sg.Frame('Dispersión', [
                    [sg.Text('Rango:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['range']:.4f} V", font=('Courier New', 10))],
                    [sg.Text('Desv. Estándar:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['std_dev']:.4f} V", font=('Courier New', 10))],
                    [sg.Text('Varianza:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['variance']:.4f} V²", font=('Courier New', 10))]
                ], border_width=0)
            ],
            [
                sg.Frame('Otros', [
                    [sg.Text('Coef. Variación:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['coefficient_variation']:.4f}", font=('Courier New', 10))],
                    [sg.Text('Semi-Intercuartil:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['semi_interquartil']:.4f} V", font=('Courier New', 10))],
                    [sg.Text('Desv. Promedio:', font=FONT_BODY), sg.Push(), sg.Text(f"{stats['mad']:.4f} V", font=('Courier New', 10))]
                ], border_width=0)
            ]
        ], element_justification='left')]
    ]

    tab_layout = [
        [
            sg.TabGroup([
                [
                    sg.Tab('Datos', [
                        [sg.Text('Lecturas de Voltaje', font=FONT_SUBTITLE, pad=(0, 10))],
                        [sg.Table(
                            values=data_table,
                            headings=headers_table,
                            auto_size_columns=False,
                            col_widths=[8, 10, 12, 12],
                            justification='right',
                            num_rows=15,
                            font=FONT_BODY,
                            expand_x=True
                        )]
                    ]),
                    sg.Tab('Estadísticas', [
                        [sg.Column(stats_layout, scrollable=True, vertical_scroll_only=True, size=(600, 300))]
                    ]),
                    sg.Tab('Gráfica', [
                        [sg.Text('Distribución de Voltajes', font=FONT_SUBTITLE, pad=(0, 10))],
                        [sg.Image(data=histogram_data, key='-HIST-', expand_x=True)]
                    ])
                ]
            ], font=FONT_BODY)
        ],
        [
            sg.Push(),
            sg.Button('Generar Reporte', font=FONT_BODY),
            sg.Button('Cerrar', font=FONT_BODY),
            sg.Push()
        ]
    ]

    layout = [
        [sg.Text('Resultados del Análisis', font=FONT_TITLE, pad=(0, 10))],
        [sg.HorizontalSeparator()],
        [sg.Column(tab_layout, pad=(0, 10), expand_x=True, expand_y=True)],
        [sg.Text('Valor Nominal: 127.00 V', font=FONT_BODY, pad=(10, 0))]
    ]

    return sg.Window('Resultados', layout, finalize=True, resizable=True)

def main():
    window = main_window()
    current_window = window

    while True:
        event, values = current_window.read()
        
        if event in (sg.WIN_CLOSED, 'Cerrar', 'Exit'):
            break
            
        if event == 'Procesar':
            file_path = values['-FILE-']
            if not file_path:
                sg.popup_error('Seleccione un archivo PDF primero')
                continue
                
            df, error = process_pdf(file_path)
            if error:
                sg.popup_error(f'Error procesando PDF:\n{error}')
                continue
                
            stats, df = calculate_statistics(df)
            histogram_data = create_histogram(df['Voltaje'])
            
            current_window.close()
            current_window = results_window(stats, df, histogram_data)
            
        if event == 'Generar Reporte PDF':
            temp_dir = tempfile.gettempdir()
            hist_path = os.path.join(temp_dir, 'histogram.png')
            with open(hist_path, 'wb') as f:
                f.write(histogram_data)
                
            report_path = os.path.join(temp_dir, 'reporte_voltaje.pdf')
            create_report(stats, hist_path, report_path)
            
            sg.popup(f'Reporte generado exitosamente:\n{report_path}')
            os.startfile(report_path)

    current_window.close()

if __name__ == '__main__':
    main()