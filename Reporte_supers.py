#-------------------------------------------
#Create By Luis Armando López
#-------------------------------------------

#We need to install some API, to create the excel
#pip install openpyxl
#pip install xlsxwritter

import pandas as pd 
try:
    super1 = input("Ingrese el nombre del archivo supers1\n")
    super2 = input("Ingrese el nombre del archivo supers2\n")
    super3 = input("Ingrese el nombre del archivo supers3\n")
    #load the file
    filename = [super1,super2, super3]
    codes = "C:\CalculoIEPS\files\template_code2.csv"
    #show all column names
    df_supers = pd.DataFrame()
    for f in filename:
        data = pd.read_csv(f)
        df_supers = df_supers.append(data)
        df_codes = pd.read_csv(codes, names = ["cve_art","type"])
    df_supers.info()
    df_codes.info()
    #in this part filter by l_imp == A
    #No iva no ieps
    df_noiva_noieps = df_supers.loc[df_supers["l_imp"] == "A"]
    df_noiva_noieps.loc['Total'] = df_noiva_noieps.select_dtypes(pd.np.number).sum()
    df_noiva_noieps.info()
    #iva no ieps
    df_iva_noieps = df_supers.loc[(df_supers["l_imp"] == "B") & (df_supers["p_imp1"] == 16)]
    df_iva_noieps.loc['Total'] = df_iva_noieps.select_dtypes(pd.np.number).sum()
    df_iva_noieps.info()
    #iva ieps
    df_iva_ieps = df_supers.loc[(df_supers["l_imp"] == "F") & (df_supers["p_imp1"] == 16) & (df_supers["p_imp2"] == 6)]
    df_iva_ieps.loc['Total'] = df_iva_ieps.select_dtypes(pd.np.number).sum()
    df_iva_ieps.info()
    #filter by D and p_imp2 == 8
    df_filter = df_supers.loc[(df_supers["l_imp"] == "D") & (df_supers["p_imp2"] == 8)]
    df_filter.info()
    #drop the rows duplicated
    df_codes1 = df_codes.drop_duplicates(subset="cve_art")
    df_codes1.info()
    #DF already clasificated by type
    df_filter_names = pd.merge(df_filter,df_codes1,left_on="cve_art", right_on="cve_art", how="left")
    df_filter_names.loc['Total'] = df_filter_names.select_dtypes(pd.np.number).sum()
    df_filter_names.info()
    #-------- filter by type
    df_botana = df_filter_names.loc[df_filter_names["type"] == "botana"]
    df_botana.loc['Total'] = df_botana.select_dtypes(pd.np.number).sum()
    df_cereal = df_filter_names.loc[df_filter_names["type"] == "cereal"]
    df_cereal.loc['Total'] = df_cereal.select_dtypes(pd.np.number).sum()
    df_chocolate = df_filter_names.loc[df_filter_names["type"] == "chocolate"]
    df_chocolate.loc['Total'] = df_chocolate.select_dtypes(pd.np.number).sum()
    df_cofiteria = df_filter_names.loc[df_filter_names["type"] == "confiteria"]
    df_cofiteria.loc['Total'] = df_cofiteria.select_dtypes(pd.np.number).sum()
    df_cavellana = df_filter_names.loc[df_filter_names["type"] == "crema de avellana"]
    df_cavellana.loc['Total'] = df_cavellana.select_dtypes(pd.np.number).sum()
    df_cavellana["TOTAL"].sum()
    df_dleche = df_filter_names.loc[df_filter_names["type"] == "dulce de leche"]
    df_dleche.loc['Total'] = df_dleche.select_dtypes(pd.np.number).sum()
    df_dverduras = df_filter_names.loc[df_filter_names["type"] == "dulce de verduras"]
    df_dverduras.loc['Total'] = df_dverduras.select_dtypes(pd.np.number).sum()
    #-----------------finish filter
    writer =  pd.ExcelWriter("Caluculo_super.xlsx", engine="xlsxwriter")
    df_noiva_noieps.to_excel(writer, sheet_name="noiva_noieps")
    df_iva_noieps.to_excel(writer, sheet_name="iva_noieps")
    df_iva_ieps.to_excel(writer, sheet_name="iva_ieps")
    df_filter_names.to_excel(writer, sheet_name="D_impuesto8_nofiltro")
    df_botana.to_excel(writer, sheet_name="Botanas")
    df_cereal.to_excel(writer, sheet_name="Cereales")
    df_chocolate.to_excel(writer, sheet_name="Chocolates")
    df_cofiteria.to_excel(writer, sheet_name="Cofiteria")
    df_cavellana.to_excel(writer, sheet_name="Crema_Avellana")
    df_dleche.to_excel(writer, sheet_name="Dulce_leche")
    df_dverduras.to_excel(writer, sheet_name="Dulce_verduras")
    print("---------------------------------------------------")
    print("El documento se esta guardando, por favor, espere.")
    print("              Puede tardar varios minutos     ")
    print("---------------------------------------------------")
    writer.save()

    


except:
    print("----------------------------------------------")
    print("----------------------------------------------")
    print("La ruta y/o nombre del archivo es incorrecto.")
    print("Por favor, intentelo de nuevo.")
    print("----------------------------------------------")
    print("----------------------------------------------")