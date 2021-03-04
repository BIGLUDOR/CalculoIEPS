#-------------------------------------------
#Create By Luis Armando López
#-------------------------------------------

#We need to install some API, to create the excel
#pip install openpyxl
#pip install xlsxwritter
import pandas as pd
try:
    namefile = input("Ingrese el nombre del archivo de cobranza\n")
    filename = f"C:\\CalculoIEPS\\Cobranza\\{namefile}.csv"
    headers = ["cia", "fec_factura", "num_factura", "serfis", "folfis", "UUID", "clave", "nombre_cliente", "RFC", "concepto", 
           "importe_fac", "base0", "base16", "iva_fac", "adeudo", "monto_pago", "iva_pag", "factor", "fecha_pago", "forma_pago", "banco", "cheque"]
    df = pd.read_csv(filename, names = headers)
    df.info()
    #Eliminate taxes duplicates
    #this funtion is used for find  the date type of the followed columns
    if df["importe_fac"].dtypes != float:
        print("No es dato flotante importe_fac")
    elif df["monto_pago"].dtypes != float:
        print("No es un dato flotante monto_pag")
    elif df["iva_fac"].dtypes != float:
        print("No es un dato flotante iva_pag")
    else:
        print("Todos los tipos de datos estan correctos")
    duplicados = df.duplicated(subset=["num_factura"])
    duplicados.value_counts()
    #to calculate th iva_pa_det and factor_det
    df["factor"] = df["monto_pago"]/df["importe_fac"]
    df["iva_pag"] = df["iva_fac"]*df["factor"]
    #doing the add of monto_pago
    sum_monto_pago = df[["num_factura","monto_pago"]].groupby(by="num_factura", as_index= False).sum()
    sum_monto_pago.info()
    #doing the add of iva_pag
    sum_iva_pag = df[["num_factura","iva_pag"]].groupby(by="num_factura", as_index= False).sum()
    sum_iva_pag.info()
    #drop the rows duplicates
    df1 = df.drop_duplicates(subset=["num_factura"])
    df1.info()
    #filter by columns we need see
    df2 = df1[["num_factura", "importe_fac"]].sort_values(by='num_factura', ascending=True)
    #join the monto_pago into DF
    merge_monto_pago = pd.merge(left=df2, right=sum_monto_pago, left_on='num_factura', right_on='num_factura')
    merge_monto_pago.shape
    #join the iva_pag into DF
    merge_iva_pag = pd.merge(left=merge_monto_pago, right=sum_iva_pag, left_on='num_factura', right_on="num_factura")
    merge_iva_pag.shape
    #Change the names
    merge_iva_pag["porcentaje"] = merge_iva_pag["monto_pago"]/merge_iva_pag["importe_fac"]
    cobro_factor = merge_iva_pag
    porcentaje = cobro_factor[["num_factura", "porcentaje"]]
    cobro_factor.loc['Total'] = cobro_factor.select_dtypes(pd.np.number).sum()
    cobro_factor.info()
    #porcentaje.head()
    #print("Match with others XLSX") 
    #atch_num_fac.head()
    #MONTHS IN THE 2019
    try:
        january_19 = pd.read_csv(r'C:\CalculoIEPS\2019\ENERO.csv', dtype={'FECHA': "string"})
        febraury_19 = pd.read_csv(r'C:\CalculoIEPS\2019\FEBRERO.csv', dtype={'FECHA': "string"})
        march_19 = pd.read_csv(r'C:\CalculoIEPS\2019\MARZO.csv', dtype={'FECHA': "string"})
        april_19 = pd.read_csv(r'C:\CalculoIEPS\2019\ABRIL.csv', dtype={'FECHA': "string"})
        may_19 = pd.read_csv(r'C:\CalculoIEPS\2019\MAYO.csv', dtype={'FECHA': "string"})
        june_19 = pd.read_csv(r'C:\CalculoIEPS\2019\JUNIO.csv', dtype={'FECHA': "string"})
        july_19 = pd.read_csv(r'C:\CalculoIEPS\2019\JULIO.csv', dtype={'FECHA': "string"})
        octuber_19 = pd.read_csv(r'C:\CalculoIEPS\2019\OCTUBRE.csv',dtype={'FECHA': "string"})
        september_19 = pd.read_csv(r'C:\CalculoIEPS\2019\SEPTIEMBRE.csv', dtype={'FECHA': "string"})
        august_19 = pd.read_csv(r'C:\CalculoIEPS\2019\AGOSTO.csv', dtype={'FECHA': "string"})
        november_19 = pd.read_csv(r'C:\CalculoIEPS\2019\NOVIEMBRE.csv', dtype={'FECHA': "string"})
        december_19 = pd.read_csv(r'C:\CalculoIEPS\2019\DICIEMBRE.csv', dtype={'FECHA': "string"})
        df2019 = pd.concat([january_19, febraury_19, march_19, april_19, may_19, june_19, july_19, octuber_19, september_19, august_19, november_19, december_19], sort=False)
    except OSError as err:
        print("No se pudo leer los exceles del 2019")
        print(sys.exc_info()[0])
    #df2019.info()
    #Months in the 2020
    try:
        january_20 = pd.read_csv(r'C:\CalculoIEPS\2020\ENERO.csv', dtype={'FECHA': "string"})
        febraury_20 = pd.read_csv(r'C:\CalculoIEPS\2020\FEBRERO.csv',dtype={'FECHA': "string"})
        march_20 = pd.read_csv(r'C:\CalculoIEPS\2020\MARZO.csv', dtype={'FECHA': "string"})
        april_20 = pd.read_csv(r'C:\CalculoIEPS\2020\ABRIL.csv', dtype={'FECHA': "string"})
        may_20 = pd.read_csv(r'C:\CalculoIEPS\2020\MAYO.csv', dtype={'FECHA': "string"})
        june_20 = pd.read_csv(r'C:\CalculoIEPS\2020\JUNIO.csv', dtype={'FECHA': "string"})
        july_20 = pd.read_csv(r'C:\CalculoIEPS\2020\JULIO.csv', dtype={'FECHA': "string"})
        octuber_20 = pd.read_csv(r'C:\CalculoIEPS\2020\OCTUBRE.csv', dtype={'FECHA': "string"})
        september_20 = pd.read_csv(r'C:\CalculoIEPS\2020\SEPTIEMBRE.csv', dtype={'FECHA': "string"})
        august_20 = pd.read_csv(r'C:\CalculoIEPS\2020\AGOSTO.csv', dtype={'FECHA': "string"})
        #november_20 = pd.read_csv("2020/NOVIEMBRE.csv")
        #decemeber_20 = pd.read_csv("2020/DICIEMBRE.csv")
        df2020 = pd.concat([january_20, febraury_20, march_20, april_20, may_20, june_20, july_20, octuber_20, september_20, august_20], sort=False)
    except OSError as err:
        print("no se pudieron leer los exceles del 2020")
        print(sys.exc_info()[0])
    try: 
        january_21 = pd.read_csv(r'C:\CalculoIEPS\2021\ENERO.csv', dtype={'FECHA': "string"})
        df2021 = pd.concat([january_21], sort=False)
    except OSError as err:
        print("No se pudieron leer los exceles del 2021")
        print(sys.exc_info()[0])
    dfmatch = pd.concat([df2019, df2020, df2021], sort=False)
    dfmatch
    dfmatch = dfmatch.sort_values(by="fac", ascending=False)
    #Part 4 sells
    fac_vts = pd.merge(left=dfmatch,right=porcentaje, left_on="fac", right_on="num_factura")
    fac_vts.shape
    fac_vts.info()
    fac_vts["importe"] = fac_vts["total"]*fac_vts["porcentaje"]
    fac_vts.head(10)
    fac_vts["IVAdeterminado"] = fac_vts["IVA"]*fac_vts["porcentaje"]
    fac_vts.head(10)
    fac_vts["IEPSdeterminado"] = fac_vts["IEPS"]*fac_vts["porcentaje"]
    fac_vts.head(10)
    fac_vts["TotalDeterminado"] = fac_vts["importe"]+fac_vts["IVAdeterminado"]+fac_vts["IEPSdeterminado"]
    fac_vts.loc['Total'] = fac_vts.select_dtypes(pd.np.number).sum()
    fac_vts.info()
    #filter to No IVa and NO IEPS
    fil_noiva_noieps = fac_vts.loc[fac_vts["l_imp"] == "A"]
    fil_noiva_noieps.loc['Total'] = fil_noiva_noieps.select_dtypes(pd.np.number).sum()
    #filter to IVA and no IEPS
    fil_iva_noieps = fac_vts.loc[fac_vts["l_imp"] == "B"]
    fil_iva_noieps.loc['Total'] = fil_iva_noieps.select_dtypes(pd.np.number).sum()
    #fil_iva_noieps.info()
    #Create filter to IVA and IEPS
    fil_iva_ieps = fac_vts.loc[fac_vts["l_imp"] == "F"]
    fil_iva_ieps.loc['Total'] = fil_iva_ieps.select_dtypes(pd.np.number).sum()
    #fil_iva_ieps.info()
    #Create filter to botanas
    fil_botanas = fac_vts.loc[fac_vts["art_sat"] == 50192100]
    fil_botanas.loc['Total'] = fil_botanas.select_dtypes(pd.np.number).sum()
    #fil_botanas
    #Create filter to chocolate
    fil_chocolate = fac_vts.loc[(fac_vts["art_sat"] >= 50161500) & (fac_vts["art_sat"] < 50161900) & (fac_vts["l_imp"] == "D")]
    fil_chocolate.loc['Total'] = fil_chocolate.select_dtypes(pd.np.number).sum()
    #fil_chocolate.info()
    #create the filter of coffee
    fil_coffee = fac_vts[(fac_vts["art_sat"] >= 50181900) & (fac_vts["art_sat"] <= 50201800) & (fac_vts["l_imp"] == "D")]
    fil_coffee.loc['Total'] = fil_coffee.select_dtypes(pd.np.number).sum()
    #fac_cobrado = pd.concat([fil_noiva_noieps, fil_iva_noieps, fil_iva_ieps, fil_botanas, fil_chocolate, fil_coffee], keys=["No IVA No IEPS", "IVA no IEPS", "IVA IEPS", 
               #                                                                                                             "botanas", "chocolate", "coffeteria"])
    #fac_cobrado.to_excel("cobranza_terminada.xlsx", index=False)
    #save the df in diferent sheet_name
    writer =  pd.ExcelWriter(f"{namefile}.xlsx", engine="xlsxwriter")
    cobro_factor.to_excel(writer, sheet_name="calculos_factores")
    fac_vts.to_excel(writer, sheet_name="All sin filtros")
    fil_noiva_noieps.to_excel(writer, sheet_name="No IVA No IEPS")
    fil_iva_noieps.to_excel(writer, sheet_name="IVA NO IEPS")
    fil_iva_ieps.to_excel(writer, sheet_name="IVA IEPS")
    fil_botanas.to_excel(writer, sheet_name="botanas")
    fil_coffee.to_excel(writer, sheet_name="coffeteria")
    print("---------------------------------------------------")
    print("El documento se esta guardando, por favor, espere.")
    print("---------------------------------------------------")
    writer.save()

except OSError as err:
    print("----------------------------------------------")
    print("----------------------------------------------")
    print("La ruta y/o nombre del archivo es incorrecto.")
    print("Por favor, intentelo de nuevo.")
    print(sys.exc_info()[0])
    print("----------------------------------------------")
    print("----------------------------------------------")