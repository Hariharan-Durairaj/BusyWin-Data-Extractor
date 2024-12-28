import os
import pandas as pd
import pyodbc
import warnings
warnings.filterwarnings("ignore")

class busywin_transactions():

    def __init__(self, busywin_folder="busywin_data", export_file_name="busywin_transactions"):
        self.busywin_folder = busywin_folder
        self.password= "ILoveMyINDIA"
        self.collect=pd.DataFrame()
        self.export_file_name= export_file_name

    def check_for_folder(self):
        if not os.path.exists(self.busywin_folder):
            # Create the folder
            os.makedirs(self.busywin_folder)
            raise Exception(f"Folder '{self.busywin_folder}' does not exist. Creating folder... Move your busywin (.bds) file to the folder ")

    def path(self):
        bw_file_names=os.listdir(self.busywin_folder)
        data_paths=[os.path.join(self.busywin_folder,i) for i in bw_file_names if i.endswith('.bds')]
        return data_paths
    
    def access_busywin_database(self, path):
        
        conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};PWD={self.password}"
        conn = pyodbc.connect(conn_str)
    
        return conn
    
    def get_transactions(self,conn):
        query="SELECT VchCode,CreationTime FROM Tran1"
        tran1=pd.read_sql(query,conn)

        query="SELECT RecType,VchType,VchCode,MasterCode1,MasterCode2,SrNo,Date,VchNo,Value1,Value2,Value3,D1 as Qty,D2,D3,D4,D5,D6,D9,D11,D12,D14,D18 FROM Tran2"
        tran2=pd.read_sql(query,conn)

        query="SELECT VchCode,ItemSrNo,HSNCode,TaxCatCode, TaxableAmt, TaxRate, TaxRate1, TaxAmt, TaxAmt1, SurchargeRate, SurchargeAmt, ActualSaleAmt FROM VchGSTSumItemWise"
        hsn=pd.read_sql(query,conn)

        query="SELECT Code as MasterCode1,MasterType, Name, ParentGrp,Alias,OldIdentity,CM1,D2 as mmrp, D3 as mlistprice,D4 as mpurcprice, D16 as mdiscountper,CreationTime,ModificationTime FROM Master1"
        masters=pd.read_sql(query,conn)

        return tran1, tran2, hsn, masters
    
    def process_transactions(self, tran1, tran2, hsn, masters):
        tran1["time"]=tran1.CreationTime.dt.time
        tran1=tran1[["VchCode","time"]]
        
        tran2.loc[tran2.RecType==2,'ItemSrNo']=tran2.loc[:,'SrNo']
        tran2.VchNo=tran2.VchNo.str.strip()

        parentgrp=masters[["MasterCode1","Name"]].rename(columns={"MasterCode1":"ParentGrp","Name":"parentgrp"})
        unit=masters[["MasterCode1","Name"]].rename(columns={"MasterCode1":"CM1","Name":"unit"})
        vchtype={"VchType":[2,9,14,15,16,19],
            "vchname":["purchase","sales","receipt","contra","journal","payment"]}
        vchtype=pd.DataFrame(vchtype)
        df=pd.merge(tran1,tran2,on="VchCode",how="right")
        df=pd.merge(df,hsn,on=["VchCode","ItemSrNo"],how="left")
        df=pd.merge(df,masters, on="MasterCode1", how="left")
        df=pd.merge(df,parentgrp,how="left",on="ParentGrp")
        df=pd.merge(df,unit,how="left",on="CM1")
        df=pd.merge(df,df[(df.RecType==1) & (df.SrNo==1)][["VchCode","Name"]].rename(columns={"Name":"Party"}),how="left",on="VchCode")
        df=pd.merge(df,vchtype,how="left",on="VchType")
        date=[]
        def Datetime(df):
            date.append(str(df["Date"])[:10]+" "+str(df["time"]))
        df.T.apply(lambda x: Datetime(x))
        df["date"]=pd.to_datetime(date)

        df=df.sort_values(by=["VchCode","RecType","SrNo"]).reset_index(drop=True)

        df["taxperwocess"]=df["TaxRate"]+df["TaxRate1"]
        df["taxper"]=df["TaxRate"]+df["TaxRate1"]+df["SurchargeRate"]
        df.loc[df.RecType==3,'Value1']=df.loc[df.RecType==3,'Value3']# check
        df["amount"]=df["D5"]-df["D11"]
        df=df[["VchType","vchname","VchCode","VchNo","Party","RecType","SrNo","date","MasterCode1","MasterType","Value1","Name","parentgrp","HSNCode","Qty","unit","D18","D4","D2","amount","D5","TaxRate","TaxAmt","TaxRate1","TaxAmt1","SurchargeRate","SurchargeAmt","taxperwocess","D14","taxper","D11","D9"]]
        df.columns=["vchtype","vchname","vchcode","vchno","party","rectype","sno","date","code","codetype","value","name","parentgrp","hsn","qty","unit","mrp","listprice","discountper","price","amount","netamount","cgstper","cgst","sgstper","sgst","cessper","cess","taxperwocess","taxwocess","taxper","tax"]
        df.loc[df.mrp==0,'mrp']=df.loc[df.mrp==0,'listprice']
        return df
        
    def close_database_connection(self,conn):
        conn.close()

    def export_data(self):
        self.collect.to_csv(self.export_file_name+".csv",index=False)

    def get_data(self):
        self.check_for_folder()
        data_paths=self.path()
        for path in data_paths:
            print(f"Connecting database {path}")
            conn=self.access_busywin_database(path)
            print("Extracting transactions...")
            tran1, tran2, hsn, masters = self.get_transactions(conn)
            self.close_database_connection(conn)
            print("Processing transactions...")
            df = self.process_transactions(tran1, tran2, hsn, masters)
            self.collect=pd.concat([self.collect,df],ignore_index=True)
            print("Completed!!!")
        return self.collect
    