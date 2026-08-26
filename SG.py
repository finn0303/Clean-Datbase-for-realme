import pandas as pd



Order = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\Order.xlsx",dtype={"IMEI": str})
TransferRaw = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\Transfer.xlsx",dtype={"IMEI": str})
Info = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\Info.xlsx",dtype={"IMEI/SN": str})
PSM = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\PSM Final July.xlsx",sheet_name="Data Store")
PSMDirect = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\PSM Final July.xlsx",sheet_name="โค้ดรับตรง")
Client = pd.read_excel(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\Client Data - Final July.xlsx")
print("Read file")
####################################################################################################################################

TransferSO = TransferRaw.sort_values(by=["transferTime"], ascending=[False])

TransferSO["Remark Type"] = (
    TransferSO["Remark"]
    .str.extract(r"([A-Za-z]+)$")
)
TransferSO["Remark SBA"] = (
    TransferSO["Remark"].str.split("-", n=2).str[1])
print("Remark1 Completed")

Remark = TransferSO[
    (
        ~TransferSO["Remark Type"].isin([
        "Transfer",
        "EX",
        "Exchage",
        "Gift",
        "Gitf",
        "Return",
        "return"
        ])
    )
]
print("Remark2 Completed")
SBASO = Remark[
    (Remark["Remark SBA"] =="SBA")
].copy()
print("TransferSO Completed")

SBASO.rename(
    columns={
        "toDistributorCode" : "Client Code",
        "transferTime" : "Date"
    },
    inplace=True
)
SBASO.drop_duplicates(subset=["IMEI"], inplace=True, keep="first")
print("TransferSO Completed")
####################################################################################################################################
OrderSO = Order.sort_values(by=["Shipped Time"], ascending=[False])
OrderSO = OrderSO[OrderSO["Return Time"].isna()]
OrderSO = OrderSO[
    OrderSO["Warehouse Name"].isin([
        "Thailand Kerry Warehouse",
        "Thailand SF warehouse"
    ])
]
OrderSO.rename(
    columns={
        "Shipped Time" : "Date",
        "Customer Code" : "Client Code"},
    inplace=True
)
OrderSO.drop_duplicates(subset=["IMEI"], inplace=True, keep="first")
print("OrderSO Completed")
####################################################################################################################################

SBASO = SBASO[["IMEI","Client Code","Date"]]
OrderSO = OrderSO[["IMEI","Client Code","Date"]]
OrderforSO = pd.concat(
    [SBASO, OrderSO],
    ignore_index=True
)
OrderforSO.sort_values(by=["Date"], ascending=[False], inplace=True)
OrderforSO.drop_duplicates(subset=["IMEI"], inplace=True, keep="first")
print("OrderforSO Completed")
####################################################################################################################################

client_mapping = {
    "DBKH8368": "DBDJ2587",
    "DBHF7717": "DBCS7452",
    "DBAU4214": "DBSD0055",
    "DBBH7146": "DBUH6450",
    "DBBC9512": "DBZR6736",
    "DBQX6987": "DBBD1522",
    "DBFC8118": "DBBH6308",
    "DBHG3689": "DBGX2217"
}


for df in [OrderforSO]:
    df["Client Code"] = (
        df["Client Code"]
        .replace(client_mapping)
    )
Info.rename(
    columns={
        "IMEI/SN" : "IMEI",
        "Store Code" : "Store ID"},
    inplace=True
)

so_cond = Info.merge(
    OrderforSO
           [
               [
                   "IMEI",
                   "Client Code"
               ]
           ],
           on="IMEI",
           how="left"
    )
Client_copy = Client.copy()

Client_copy["Client Code for Map"] = Client_copy["Client Code"]
Client_copy["Channel for Map"] = Client_copy["Client Channel"]
so_cond = so_cond.merge(
    Client_copy[
        [
            "Client Code",
            "Client Code for Map",
            "Client Name (GS-System)",
            "Client Type",
            "Client Subtype",
            "Client Channel",
            "Channel for Map",
            "Status",
            "Main Zone",
            "Main RSD"
        ]
    ],
    on="Client Code",
    how="left"
)
PSM["so_rsd"] = (
    PSM["RSD"].str.split("-", n=1).str[0])

PSM["so_rsm"] = (
    PSM["RSM"].str.split("-", n=1).str[0]
)


so_cond = so_cond.merge(
    PSM[
        ["Store ID",
         "BU",
         "DB ID",
         "Responsible",
         "DB ID2",
         "Province",
         "Area",
         "Area Code",
         "District",
         "Big Type",
         "Store Type",
         "Subclass2",
         "Zone",
         "so_rsd",
         "so_rsm"
         ]
    ],
    on="Store ID",
    how="left"
)

Client["Client Channel for map"] = Client["Client Channel"].copy()

Client_DB1 = Client[
    [
        "Client Code",
        "Client Name (GS-System)",
        "Client Channel",
        "Client Channel for map",
    ]
].rename(
    columns={
        "Client Code" : "DB1",
        "Client Name (GS-System)": "DB1 name",
        "Client Channel": "DB1 channel",
        "Client Channel for map": "DB1 channel for map",
    }
)

Client_DB2 = Client[
    [
        "Client Code",
        "Client Name (GS-System)",
        "Client Channel",
        "Client Channel for map",
    ]
].rename(
    columns={
        "Client Code" : "DB2",
        "Client Name (GS-System)": "DB2 name",
        "Client Channel": "DB2 channel",
        "Client Channel for map": "DB2 channel for map",
    }
)
so_cond = so_cond.merge(
    Client_DB1,
    left_on="DB ID",
    right_on="DB1",
    how="left"
).drop(columns="DB1")

so_cond = so_cond.merge(
    Client_DB2,
    left_on="DB ID2",
    right_on="DB2",
    how="left"
).drop(columns="DB2")

Channel_mapping = {
    "Hub" : "GT_Hub",
    "KD" : "GT_KD&Dealer",
    "Dealer" : "GT_KD&Dealer",
    "Brand shop_GT" : "GT_Brand shop_GT",
    "IT_CITY" : "KA",
    "JAYMART" : "KA",
    "Advice" : "KA",
    "TG" : "KA",
    "JIB" : "KA",
    "SINGER" : "KA",
    "MAXLINK" : "KA",
    "COM_7" : "Online_Com_7",
    "OPT_AIS" : "OPT_AIS",
    "OPT_TRUE" : "OPT-NOT AIS",
    "OPT_DTAC" : "OPT-NOT AIS",
    "OPT_CJ" : "OPT-NOT AIS",
    "After-Service" : "After-Service",
    "Others" : "Others",
    "Overseas" : "Overseas",
    "TH-realme" : "TH-realme"
}

so_cond["Channel for Map"] = so_cond["Channel for Map"].replace(Channel_mapping)

so_cond["SPU_Spec_Full"] = (
    so_cond["SPU Name"]
    .str.cat(so_cond["SKU Spec"], sep=" ")
)

so_cond["SPU_Spec_abbr"] = None
so_cond["spec_tier"] = None
so_cond["E_date"] = None
so_cond["ST Client"] = None
so_cond["ST Client Name"] = None
so_cond["ST Client Channel"] = None
so_cond["Store id by client"] = None
so_cond["Store No Pc 15 Days"] = None
####################################################################################################################################
Store_Direct = PSMDirect[[
    "Store ID",
    "Client ID สั่งตรง",
    "Client Name"]].rename(columns=
    {"Store ID" : "Store ID",
     "Client ID สั่งตรง" : "Client_Direct ID",
     "Client Name" : "Client_Direct Name"})
so_cond = so_cond.merge(
    Store_Direct,
    on="Store ID",
    how="left"
)

SG = so_cond[
    [
        "IMEI",
        "Order Type",
        "Product Type",
        "SKU ID",
        "Series Name",
        "spec_tier",
        "SPU Name",
        "SKU Spec",
        "SPU_Spec_Full",
        "SPU_Spec_abbr",
        "Activate Time",
        "Client Code",
        "Client Code for Map",
        "Client Name (GS-System)",
        "Client Type",
        "Client Subtype",
        "Client Channel",
        "Channel for Map",
        "Status",
        "Main Zone",
        "Main RSD",
        "ST Client",
        "ST Client Name",
        "ST Client Channel",
        "Store ID",
        "Store id by client",
        "Store Name",
        "BU",
        "DB ID",
        "DB1 name",
        "DB1 channel",
        "DB1 channel for map",
        "Responsible",
        "DB ID2",
        "DB2 name",
        "DB2 channel",
        "Province",
        "Area",
        "Area Code",
        "District",
        "Big Type",
        "Store Type",
        "Subclass2",
        "Client_Direct ID",
        "Client_Direct Name",
        "Store No Pc 15 Days",
        "Zone",
        "so_rsd",
        "so_rsm"

    ]]

with (pd.ExcelWriter(r"C:\Users\RM-NB-002\Desktop\CN\CN GT\Python\SG\2026\July\Result.xlsx",engine="openpyxl") as writer):
    OrderforSO.to_excel(writer,sheet_name="OrderforSO",index=False)
    SBASO.to_excel(writer,sheet_name="TransferSO",index=False)
    OrderSO.to_excel(writer,sheet_name="OrderSO",index=False)
    SG.to_excel(writer,sheet_name="SG",index=False)

print("run completed")