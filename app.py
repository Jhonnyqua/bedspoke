"border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        # Formato de celdas y bandas alternas
        cell_fmt = wb.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        band_fmt = wb.add_format({
            "border": 1,
            "bg_color": "#F2F2F2",
            "align": "left",
            "valign": "vcenter",
        })

        # Escribir encabezados
        for col, name in enumerate(grouped.columns):
            ws.write(0, col, name, header_fmt)

        # Ajustar anchos de columna
        ws.set_column("A:A", 35, cell_fmt)  # Dirección
        ws.set_column("B:B", 25, cell_fmt)  # Encargado
        ws.set_column("C:C", 40, cell_fmt)  # Llave M

        # Aplicar bandas alternas
        for row in range(1, len(grouped) + 1):
            ws.set_row(row, None, band_fmt if row % 2 == 0 else cell_fmt)

    output.seek(0)
    return grouped, output.read()

# ----------------------------------------------------------
# 4. Interfaz Streamlit
# ----------------------------------------------------------
st.set_page_config(page_title="Reporte Llaves M", layout="wide")

st.title("🗝️ Reporte de Llaves M (ordenado por Cleaner)")

csv_file = st.file_uploader("📥 Sube tu CSV de IGMS", type="csv")
if csv_file:
    df_grp, excel_bytes = create_grouped_excel(csv_file)
    st.success("✅ Reporte generado correctamente")
    st.dataframe(df_grp, use_container_width=True)
    st.download_button(
        label="⬇️ Descargar Reporte (Excel)",
        data=excel_bytes,
        file_name="reporte_llaves_m.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("📄 Esperando que subas un archivo CSV de IGMS")
