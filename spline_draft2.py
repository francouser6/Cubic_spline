# -*- coding: utf-8 -*-
"""
Draft 1 calculando cubic spline con scipy
"""
import numpy as np
import pandas as pd
from io import BytesIO
import streamlit as st

import subprocess
subprocess.run(["pip", "install", "openpyxl"])


# Configuración de la página
st.set_page_config(page_title="Financial Risk Management", page_icon="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS14bSWA3akUYXe-VV04Nw2K0QnQCwCV9SG8g&s")
st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS14bSWA3akUYXe-VV04Nw2K0QnQCwCV9SG8g&s", width=250)
st.title("Interpolación Cúbica")
st.markdown("### Plataforma de estimación de interpolación cúbica para múltiples datos y archivos")

def cubic_spline(df, input_column_name, output_column_names, points):
    if input_column_name not in df.columns:
        return "¡Algo está mal! ¡El nombre de la columna de entrada no existe en el DataFrame!"

    input_column = df[input_column_name].dropna().values
    output_columns = output_column_names

    n = len(input_column)
    if n <= 1:
        return "¡Algo está mal! ¡La columna de entrada debe tener más de un valor!"

    # Ordenar los valores de xin y yin
    sorted_indices = np.argsort(input_column)
    xin = input_column[sorted_indices]
    results_dict = {'Puntos Evaluados': points}

    for output_column_name in output_columns:
        if output_column_name not in df.columns:
            return f"¡Algo está mal! ¡El nombre de la columna de salida '{output_column_name}' no existe en el DataFrame!"

        yin = df[output_column_name].dropna().values
        if len(yin) != n:
            return f"¡Algo está mal! ¡La columna de salida '{output_column_name}' no tiene la misma longitud que la columna de entrada!"

        # Ordenar yin de acuerdo a los índices de xin
        yin = yin[sorted_indices]

        yt = np.zeros(n)
        u = np.zeros(n - 1)

        yt[0] = 0
        u[0] = 0

        for i in range(1, n - 1):
            sig = (xin[i] - xin[i - 1]) / (xin[i + 1] - xin[i - 1])
            p = sig * yt[i - 1] + 2
            yt[i] = (sig - 1) / p
            u[i] = (yin[i + 1] - yin[i]) / (xin[i + 1] - xin[i]) - (yin[i] - yin[i - 1]) / (xin[i] - xin[i - 1])
            u[i] = (6 * u[i] / (xin[i + 1] - xin[i - 1]) - sig * u[i - 1]) / p

        qn = 0
        un = 0
        yt[n - 1] = (un - qn * u[n - 2]) / (qn * yt[n - 2] + 1)

        for k in range(n - 2, 0, -1):
            yt[k] = yt[k] * yt[k + 1] + u[k]

        def evaluate_point(x):
            klo = 0
            khi = n - 1
            while khi - klo > 1:
                k = (khi + klo) // 2
                if xin[k] > x:
                    khi = k
                else:
                    klo = k

            h = xin[khi] - xin[klo]
            a = (xin[khi] - x) / h
            b = (x - xin[klo]) / h
            y = a * yin[klo] + b * yin[khi] + ((a ** 3 - a) * yt[klo] + (b ** 3 - b) * yt[khi]) * (h ** 2) / 6

            return y

        results = [evaluate_point(x) for x in points]
        results_dict[output_column_name] = results

    output_df = pd.DataFrame(results_dict)

    return output_df


def main():
    st.write("Sube un archivo y selecciona la columna de entrada y las columnas de salida para aplicar la interpolación cúbica.")

    uploaded_file = st.file_uploader("Sube tu archivo (xlsx, xls, xlsm, csv)", type=["xlsx", "xls", "xlsm", "csv"])

    if uploaded_file:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            sheets = {'Sheet1': df}
        else:
            sheets = pd.read_excel(uploaded_file, sheet_name=None)

        st.write("Hojas disponibles:", list(sheets.keys()))

        exclude_sheets = st.multiselect("Selecciona las hojas que deseas excluir del cálculo", list(sheets.keys()))
        sheets = {sheet_name: df for sheet_name, df in sheets.items() if sheet_name not in exclude_sheets}

        start_point = st.number_input("Inicio del rango de puntos a evaluar", value=0, format="%d")
        end_point = st.number_input("Fin del rango de puntos a evaluar", value=13000, format="%d")
        num_points = st.number_input("Número de puntos a evaluar", value=100, min_value=1, format="%d")

        points = np.linspace(start_point, end_point, num_points).astype(int)

        output_sheets = {}
        for sheet_name, df in sheets.items():
            st.write(f"Configuración para la hoja: {sheet_name}")
            header_row = st.number_input(f"Fila del encabezado para la hoja {sheet_name}", min_value=0, max_value=len(df)-1, value=0)
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:].dropna(how='all').apply(pd.to_numeric, errors='coerce')

            if sheet_name == "391 - EUR EIOPA UFR Curve":
                df.columns = df.columns.str.replace(r'(\d+)\s*yr', lambda x: x.group(1), regex=True)

            input_column_name = st.selectbox(f"Selecciona la columna de entrada para la hoja {sheet_name}", df.columns)
            output_column_names = st.multiselect(f"Selecciona las columnas de salida para la hoja {sheet_name}", df.columns.difference([input_column_name]))

            output_df = cubic_spline(df, input_column_name, output_column_names, points)
            if isinstance(output_df, pd.DataFrame):
                output_sheets[sheet_name] = output_df
            else:
                st.error(output_df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, output_df in output_sheets.items():
                output_df.to_excel(writer, sheet_name=sheet_name, index=False)

        st.download_button(
            label="Descargar resultados",
            data=output.getvalue(),
            file_name="resultados_interpolacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write("Visualización de resultados:")
        for sheet_name, output_df in output_sheets.items():
            st.write(f"Resultados para la hoja: {sheet_name}")
            st.dataframe(output_df)

        # Añadir previsualización de las primeras 5 filas y 20 columnas de cada hoja
        st.write("Previsualización de las primeras 5 filas de cada hoja:")
        for sheet_name, df in sheets.items():
            if st.button(f"Mostrar las primeras 5 filas de {sheet_name}"):
                st.write(f"Primeras 5 filas de la hoja: {sheet_name}")
                preview_df = df.iloc[:5, :20].dropna(axis=1, how='all')
                st.dataframe(preview_df)

if __name__ == "__main__":
    main()
