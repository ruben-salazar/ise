import pandas as pd
import os
import sys


def csv_a_excel_matriz(archivo_csv, archivo_excel):
    """
    Convierte CSV largo a matriz Excel.

    Campo1 -> filas únicas
    Campo2 -> columnas únicas
    Campo3 -> valor intersección
    Campo4 -> ignorado

    EMPTY -> celda vacía
    """

    # Leer CSV ignorando primera línea
    df = pd.read_csv(
        archivo_csv,
        skiprows=1,
        header=None,
        names=["campo1", "campo2", "campo3", "campo4"],
        dtype=str
    )

    # limpiar espacios
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

    # convertir EMPTY a vacío
    df["campo3"] = df["campo3"].replace("EMPTY", "")

    # eliminar cuarta columna
    df = df[["campo1", "campo2", "campo3"]]

    # pivotear
    matriz = df.pivot(
        index="campo1",
        columns="campo2",
        values="campo3"
    )

    # vacíos reales
    matriz = matriz.fillna("")

    # reset index para que primera columna sea visible
    matriz = matriz.reset_index()

    # nombre vacío de esquina superior izquierda
    matriz.columns.name = None

    # exportar
    matriz.to_excel(
        archivo_excel,
        index=False,
        engine="openpyxl"
    )

    print(f"Archivo generado correctamente: {archivo_excel}")
    print(f"Filas únicas: {len(df['campo1'].unique())}")
    print(f"Columnas únicas: {len(df['campo2'].unique())}")


if __name__ == "__main__":

    if len(sys.argv) != 3:
        print("Uso:")
        print("python csv_to_excel.py entrada.csv salida.xlsx")
        sys.exit(1)

    archivo_csv = sys.argv[1]
    archivo_excel = sys.argv[2]

    if not os.path.exists(archivo_csv):
        print(f"No existe el archivo: {archivo_csv}")
        sys.exit(1)

    csv_a_excel_matriz(
        archivo_csv,
        archivo_excel
    )
