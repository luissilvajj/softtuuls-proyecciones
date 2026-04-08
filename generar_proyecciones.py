import pandas as pd
import xlsxwriter

def generar_excel_softtuuls(nombre_archivo='Proyecciones_Softtuuls.xlsx'):
    # Variables del modelo
    meses = 12
    inversion_inicial = 10200
    costo_fijo_base = 150
    incremento_costo_mes_7 = 50
    precio_setup = 250
    precio_mensual = 40
    clientes_nuevos_por_mes = 5

    # Estructuras para guardar los datos
    datos = []
    clientes_activos = 0
    flujo_acumulado = -inversion_inicial
    
    # Fila del Mes 0 (Inversión)
    datos.append({
        'Mes': 0,
        'Nuevos Clientes': 0,
        'Clientes Activos': 0,
        'Ingreso Setup ($)': 0,
        'Ingreso Recurrente MRR ($)': 0,
        'Ingreso Total ($)': 0,
        'Costos Mensuales ($)': 0,
        'Flujo de Caja del Mes ($)': -inversion_inicial,
        'Balance Acumulado ($)': flujo_acumulado
    })

    # Proyección Mes 1 al 12
    for mes in range(1, meses + 1):
        clientes_activos += clientes_nuevos_por_mes
        ingreso_setup = clientes_nuevos_por_mes * precio_setup
        ingreso_mrr = clientes_activos * precio_mensual
        ingreso_total = ingreso_setup + ingreso_mrr
        
        costo_mensual = costo_fijo_base if mes < 7 else costo_fijo_base + incremento_costo_mes_7
        
        flujo_mes = ingreso_total - costo_mensual
        flujo_acumulado += flujo_mes
        
        datos.append({
            'Mes': mes,
            'Nuevos Clientes': clientes_nuevos_por_mes,
            'Clientes Activos': clientes_activos,
            'Ingreso Setup ($)': ingreso_setup,
            'Ingreso Recurrente MRR ($)': ingreso_mrr,
            'Ingreso Total ($)': ingreso_total,
            'Costos Mensuales ($)': costo_mensual,
            'Flujo de Caja del Mes ($)': flujo_mes,
            'Balance Acumulado ($)': flujo_acumulado
        })

    # Crear DataFrame
    df = pd.DataFrame(datos)

    # Crear archivo Excel con XlsxWriter
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Proyecciones', index=False)

    # Obtener el workbook y el worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Proyecciones']

    # --- FORMATOS VISUALES ---
    formato_moneda = workbook.add_format({'num_format': '$#,##0', 'align': 'center'})
    formato_centro = workbook.add_format({'align': 'center'})
    formato_cabecera = workbook.add_format({
        'bold': True, 
        'bg_color': '#1f4e78', 
        'font_color': 'white', 
        'align': 'center',
        'border': 1
    })

    # Aplicar formatos a las columnas
    worksheet.set_column('A:A', 5, formato_centro)
    worksheet.set_column('B:C', 15, formato_centro)
    worksheet.set_column('D:I', 20, formato_moneda)

    # Aplicar formato a la cabecera
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, formato_cabecera)

    # --- GRÁFICO 1: PUNTO DE EQUILIBRIO (Líneas) ---
    chart_roi = workbook.add_chart({'type': 'line'})
    
    # Línea de Balance Acumulado
    chart_roi.add_series({
        'name': 'Balance Acumulado',
        'categories': ['Proyecciones', 1, 0, 13, 0],
        'values': ['Proyecciones', 1, 8, 13, 8],  # columna I (índice 8) = Balance Acumulado
        'line': {'color': '#1f4e78', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#1f4e78'}}
    })

    chart_roi.set_title({'name': 'Punto de Equilibrio (Break-even) a 12 Meses'})
    chart_roi.set_x_axis({'name': 'Meses'})
    chart_roi.set_y_axis({'name': 'Balance Acumulado ($ USD)', 'major_gridlines': {'visible': True}})
    chart_roi.set_legend({'position': 'bottom'})
    chart_roi.set_size({'width': 700, 'height': 400})

    # Insertar el gráfico 1
    worksheet.insert_chart('K2', chart_roi)

    # --- GRÁFICO 2: CRECIMIENTO DE INGRESOS (Barras Apiladas) ---
    chart_ingresos = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

    # Serie MRR (columna E, índice 4)
    chart_ingresos.add_series({
        'name': 'Ingreso Recurrente (MRR)',
        'categories': ['Proyecciones', 1, 0, 13, 0],
        'values': ['Proyecciones', 1, 4, 13, 4],
        'fill': {'color': '#00b050'}  # Verde
    })
    
    # Serie Setup (columna D, índice 3)
    chart_ingresos.add_series({
        'name': 'Ingreso por Setup',
        'categories': ['Proyecciones', 1, 0, 13, 0],
        'values': ['Proyecciones', 1, 3, 13, 3],
        'fill': {'color': '#92d050'}  # Verde claro
    })

    chart_ingresos.set_title({'name': 'Composición de Ingresos Mensuales'})
    chart_ingresos.set_x_axis({'name': 'Meses'})
    chart_ingresos.set_y_axis({'name': 'Ingresos ($ USD)'})
    chart_ingresos.set_legend({'position': 'bottom'})
    chart_ingresos.set_size({'width': 700, 'height': 400})

    # Insertar el gráfico 2
    worksheet.insert_chart('K23', chart_ingresos)

    # Guardar y cerrar
    writer.close()
    print(f"✅ Archivo '{nombre_archivo}' generado con éxito. Ábrelo en Excel.")

if __name__ == "__main__":
    generar_excel_softtuuls()
