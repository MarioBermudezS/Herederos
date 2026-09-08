import pandas as pd

def app():
    print("========================================")
    print("   GENERADOR DE INFORMES - HEREDEROS    ")
    print("========================================")
    
    archivo = "BaseDatos2026.xlsx"
    ano = int(input("Introduce el año (ej. 2026): "))
    mes = input("Introduce el mes (ej. Junio, Enero...): ").capitalize()
    
    tipo = input("¿Quieres consultar una sola tienda o un conjunto? (1 = Una tienda, 2 = Varias): ")
    
    if tipo == "1":
        tienda = input("Introduce el nombre de la tienda (ej. Torremolinos): ")
        tiendas = [tienda]
    else:
        lista_tiendas = input("Introduce las tiendas separadas por comas (ej. Torremolinos, Fuengirola): ")
        tiendas = [t.strip() for t in lista_tiendas.split(",")]

    print(f"\nGenerando informe para {tiendas} ({mes} {ano})...\n")
    
    try:
        df = pd.read_excel(archivo, sheet_name='BS')
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return
        
    mask = (df['Año'] == ano) & (df['Mes'] == mes) & (df['Departamento'].isin(tiendas))
    df_filtered = df[mask]
    
    if df_filtered.empty:
        print("¡Aviso! No se han encontrado datos para esos criterios.")
        return

    resumen = df_filtered.groupby('Resultados')['Importe D'].sum().to_dict()
    
    def get_val(cat):
        return resumen.get(cat, 0.0)
    
    ventas = get_val('Ventas')
    coste_ventas = get_val('Consumo Ventas')
    margen_bruto = ventas - coste_ventas
    r_bruta = (margen_bruto / ventas) if ventas != 0 else 0.0
    
    otros_ingresos = get_val('Otros Ingresos')
    ingresos_operativos = margen_bruto + otros_ingresos
    
    gastos_personal = get_val('Gastos Personal')
    alquileres = get_val('Alquileres')
    reparaciones = get_val('Reparaciones')
    seguros = get_val('Seguros')
    suministros = get_val('Suministros')
    otros_servicios = get_val('Otros Servicios')
    
    total_gastos_operativos = alquileres + reparaciones + seguros + suministros + otros_servicios
    amortizaciones = get_val('Amortizaciones')
    
    baii = ingresos_operativos - gastos_personal - total_gastos_operativos - amortizaciones
    
    gastos_financieros = get_val('Gastos Financieros')
    ingresos_financieros = get_val('Ingresos Financieros')
    rdo_financiero = ingresos_financieros - gastos_financieros
    
    resultados_extraordinarios = get_val('Resultados Extraordinarios')
    bai = baii + rdo_financiero + resultados_extraordinarios
    
    data_out = [
        ("Ventas", ventas),
        ("Coste Ventas", coste_ventas),
        ("MARGEN BRUTO", margen_bruto),
        ("R. Bruta", r_bruta),
        ("Otros Ingresos", otros_ingresos),
        ("Ingresos Operativos", ingresos_operativos),
        ("Gastos Personal", gastos_personal),
        ("Alquileres", alquileres),
        ("Reparaciones", reparaciones),
        ("Seguros", seguros),
        ("Suministros", suministros),
        ("Otros Servicios", otros_servicios),
        ("TOTAL GASTOS OPERATIVOS", total_gastos_operativos),
        ("Amortizaciones", amortizaciones),
        ("B.A.I.I.", baii),
        ("Gastos Financieros", gastos_financieros),
        ("Ingresos Financieros", ingresos_financieros),
        ("RDO. FINANCIERO", rdo_financiero),
        ("Resultados Extraordinarios", resultados_extraordinarios),
        ("B.A.I.", bai)
    ]
    
    df_resultado = pd.DataFrame(data_out, columns=["Resultados", f"{mes} {ano}"])
    print(df_resultado.to_string(index=False))
    
    guardar = input("\n¿Quieres exportar este informe a un archivo de Excel? (s/n): ").lower()
    if guardar == 's':
        nombre_salida = f"Informe_{'_'.join(tiendas)}_{mes}_{ano}.xlsx"
        df_resultado.to_excel(nombre_salida, index=False)
        print(f"¡Informe guardado con éxito como '{nombre_salida}'!")

if __name__ == "__main__":
    app()
