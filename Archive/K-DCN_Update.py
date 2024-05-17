
import pandas as pd


DCNfrmkola_path = r'C:\Python_SPI\Global_DCN_Test\DCNfrmKola.xlsx'
DCNinJobQ_path = r'C:\Python_SPI\Global_DCN_Test\DCNinJobQ.xlsx'
DCNOrdered_path=r'C:\Python_SPI\Global_DCN_Test\DCNOrdered.xlsx'
ProjectByIH_path=r'C:\Python_SPI\Global_DCN_Test\Query_Files\ProjectByIH.xlsx'
AMObjects_path=r'C:\Python_SPI\Global_DCN_Test\Query_Files\AMObjects.xlsx'
########converting into  Dataframes########
df_kola = pd.read_excel(DCNfrmkola_path,sheet_name='DCNfrmKola')
df_jobq = pd.read_excel(DCNinJobQ_path, sheet_name='DCNinJobQ')
df_DCNOrdered=pd.read_excel(DCNOrdered_path, sheet_name='DCNOrdered')
df_ProjectByIH=pd.read_excel(ProjectByIH_path, sheet_name='ProjectByIH')
df_AMObjects=pd.read_excel(AMObjects_path, sheet_name='AMObjects')

def UpdateK_DCN():
    # Update 'DCNinJobQ' column in DCNfrmKola where the condition is True
    DCNJobQUpdate_condition = df_kola['DCNG'].isin(df_jobq['DCN Design Change Notice'])
    df_kola.loc[DCNJobQUpdate_condition, 'DCNinJobQ'] = 'Yes'
    df_kola.to_excel(DCNfrmkola_path, index=False)
    print("DCNJobQUpdate_condition",df_kola)

    # Update 'DCNOrdered' column in DCNfrmKola where the condition is True
    DCNOrderUpdate_condition=df_kola['DCNG'].isin(df_DCNOrdered['DCN Number'])
    df_kola.loc[DCNOrderUpdate_condition, 'DCNOrdered'] = 'Yes'
    df_kola.to_excel(DCNfrmkola_path, index=False)
    print("DCNOrderUpdate_condition",df_kola)

    # Update 'PPLNotSigned' column in DCNfrmKola where the combined condition is True
    condition1=df_kola['DCNG'].isin(df_jobq['DCN Design Change Notice'])
    condition2 = (df_jobq['DCN Archive Date'] == "1901-01-01")
    PPLNotSigned_condtion=condition1 & condition2
    df_kola.loc[PPLNotSigned_condtion, 'PPLNotSigned'] = 'Yes'
    df_kola.to_excel(DCNfrmkola_path, index=False)
    print("PPLNotSigned_condtion",df_kola)

    #Update 'DCNBySCP' column in DCNfrmKola where the condition is True
    sub_object_values = df_ProjectByIH['Sub Object'].unique()
    sub_object_list = sub_object_values.tolist()
    ProjectByIH_condition = df_kola['Object'].isin(sub_object_list)
    df_kola.loc[ProjectByIH_condition, 'DCNBySCP'] = 'Yes'
    print("ProjectByIH_condition",df_kola)
    df_kola.to_excel(DCNfrmkola_path, index=False)

    #Update 'AMObjects' column in DCNfrmKola where the condition is True
    AMObjects_values = df_AMObjects['AMObject'].unique()
    sub_object_list = AMObjects_values.tolist()
    AMObjects_condition = df_kola['Object'].isin(sub_object_list)
    df_kola.loc[AMObjects_condition, 'AMObjects'] = 'Yes'
    df_kola.to_excel(DCNfrmkola_path, index=False)
    print("AMObjects",df_kola)

    ##Deleting null values in Duplicate column in Kola Excel
    df_kola.dropna(subset=['Duplicate'], inplace=True)
    df_kola.to_excel(DCNfrmkola_path, index=False)
    ##Deleting null value in DCNInJOQ column in Kola Excel
    df_kola.dropna(subset=['DCNinJobQ'], inplace=True)
    df_kola.to_excel(DCNfrmkola_path, index=False)
    print("Final",df_kola)
    return()

UpdateK_DCN()

    







    





    




