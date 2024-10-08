import numpy as np
import pandas as pd

Config=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Config")
Banka_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Banka")
Masraf_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Masraf")

Bulut_All=pd.read_excel(Config.iloc[4,1], dtype={'Vergi No': str})
Bulut_All["Tarih"]=Bulut_All["Tarih"].dt.date

def assign_code(row):
    if row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] == 24 and row['İşlem Tipi'] == 'MASRAF':
        return 'TB_EFTM'
    elif row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] == 1070 and row['İşlem Tipi'] == 'MASRAF':
        return 'TB_POSM'
    elif row['Banka'] == 'AKBANK' and row['Fonksiyon Kodu 1'] == 'MSC' and row['Fonksiyon Kodu 2'] == '00F8':
        return 'AK_EFTM'
    elif row['Banka'] == 'AKBANK' and row['Fonksiyon Kodu 1'] == 'MSC' and row['Fonksiyon Kodu 2'] == '00UC':
        return 'AK_POSM'
    elif row['Banka'] == 'HALKBANK' and row['Fonksiyon Kodu 1'] == 'COC' and 'Aidat' in row['Açıklama']:
        return 'HB_POSM'
    elif row['Banka'] == 'İŞ BANKASI' and row['Fonksiyon Kodu 1'] == "KOM" and row['İşlem Tipi'] == 'MASRAF':
        return 'IS_EFTM'
    elif row['Banka'] == 'İŞ BANKASI' and row['Fonksiyon Kodu 1'] == "CCP" and row['İşlem Tipi'] == 'MASRAF':
        return 'IS_POSM'
    elif row['Banka'] == 'GARANTİ BBVA' and row['Fonksiyon Kodu 1'] == "BT19" and row['İşlem Tipi'] == 'MASRAF':
        return 'GB_EFTM'
    elif row['Banka'] == 'YAPI KREDİ' and row['Fonksiyon Kodu 1'] == "CHG" and row['İşlem Tipi'] == 'MASRAF':
        return 'YK_EFTM'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == "MSC" and row['İşlem Tipi'] == 'MASRAF' and 'POS YAZILIM' in row['Açıklama']:
        return 'QB_POSM'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == "MSC" and row['İşlem Tipi'] == 'MASRAF' and ' DBS Dönem Ücreti ' in row['Açıklama']:
        return 'QB_DBSM'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 2'] in ["UPSTPOSUCRET13", "UPSTPOSUCRET8", "UPSUYEBLOKE14"] and row['İşlem Tipi'] == 'MASRAF':
        return 'VB_POSM'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 2'] == "TKMTKOMTAHSILAT" and row['İşlem Tipi'] == 'MASRAF':
        return 'VB_BTMM'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 1'] == "CHG" and row['Fonksiyon Kodu 2'] == "FYTSMSRM02" and row['İşlem Tipi'] == 'MASRAF':
        return 'VB_EFTM'
    elif row['Banka'] == 'ZİRAAT KATILIM' and row['Fonksiyon Kodu 1'] in ["MASTT", "KTMGC"] and row['İşlem Tipi'] == 'MASRAF':
        return 'ZK_BTMM'
    elif row['Banka'] == 'ZİRAAT BANKASI' and row['Fonksiyon Kodu 1'] == "XXX" and row['Fonksiyon Kodu 2'] == 'UYESUCRT' :
        return 'ZB_POSM'
    else:
        return ''

Bulut_All['KOD'] = Bulut_All.apply(assign_code, axis=1)


Bulut_Masraf = Bulut_All[Bulut_All['KOD'] !=""]

Bulut_Code1=Bulut_Masraf.merge(Banka_Conf[["Firma IBAN","BANKA HESAP KODU", "Muhasebe Kodu"]], on="Firma IBAN", how="left")

Bulut_Code=Bulut_Code1.merge(Masraf_Conf, on="KOD")

Bulut_Code["Tutar2"]=-Bulut_Code["Tutar"]

Bulut_Code.rename(columns={"Tutar": 'Banka_T'}, inplace=True)

V3_data = pd.melt(
    Bulut_Code,
    id_vars=[col for col in Bulut_Code.columns if col not in ['Banka_T', 'Tutar2']],
    value_vars=['Banka_T', 'Tutar2'],
    var_name='Type',
    value_name='Tutar'
).sort_values(by=["İşlem Kodu", "Tutar"] )

V3_Format=V3_data[["Tarih"]].copy()
V3_Format.rename(columns={"Tarih": 'DocumentDate'}, inplace=True)

V3_Format["DocumentNumber"]=V3_data["İşlem Kodu"]
V3_Format["Description"]=V3_data["Banka"]+"-"+V3_data["SATIR AÇIKLAMASI"]
V3_Format["BankCurrAccCode"] = np.where(V3_data["Type"] == "Banka_T", V3_data["BANKA HESAP KODU"], "")
V3_Format["BankOpTypeCode"]= V3_data["BANKA İŞLEM TİPİ"]
V3_Format["DueDate"]= ""
V3_Format["LineDescription"]= V3_data["Banka"]+"-"+V3_data["SATIR AÇIKLAMASI"]
V3_Format["LineDocumentNumber"]= np.where(V3_data["Type"] ==  "Banka_T", 1, 2)
V3_Format["DocumentTypeCode"]= ""
V3_Format["DocumentTypeDescription"]= ""
V3_Format["PaymentMethod"]= "Banka"
V3_Format["GLAccCode"]= np.where(V3_data["Type"] == "Banka_T", V3_data["Muhasebe Kodu"], V3_data["HESAP KODU"])
V3_Format["CostCenterCode"]= "E1"
V3_Format["GLTypeCode"]= ""
V3_Format["ImportFileNumber"]= ""
V3_Format["ExportFileNumber"]= ""
V3_Format["DocCurrencyCode"]= "TRY"
V3_Format["ExchangeRate"]= 1
V3_Format["Debit"]= np.where(V3_data["Tutar"] > 0, V3_data["Tutar"], 0)
V3_Format["Credit"]= np.where(V3_data["Tutar"] < 0, -V3_data["Tutar"], 0)
V3_Format["ATAtt01"]= ""
V3_Format["ATAtt02"]= ""
V3_Format["ATAtt03"]= ""
V3_Format["ATAtt04"]= ""
V3_Format["ATAtt05"]= ""
V3_Format["FTAtt01"]= ""
V3_Format["FTAtt02"]= ""
V3_Format["FTAtt03"]= ""
V3_Format["FTAtt04"]= ""
V3_Format["FTAtt05"]= ""

V3_Format.to_excel(Config.iloc[12,1], index=False)