import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

excel_file = 'input_data.xlsx'
df = pd.read_excel(excel_file, skiprows=5)

df.columns = [col.strip() for col in df.columns]

def safe_str(value):
    return '' if pd.isna(value) else str(value)

def create_sales_order_xml(df):
    root = ET.Element("NMEXML", EximID="13", BranchCode="2040822216", ACCOUNTANTCOPYID="")

    transactions = ET.SubElement(root, "TRANSACTIONS", OnError="CONTINUE")
    unique_orders = df['Order No'].unique()

    for order_no in unique_orders:
        sales_order = ET.SubElement(transactions, "SALESORDER", operation="Add", REQUESTID="1")
        ET.SubElement(sales_order, "TRANSACTIONID")
        
        order_df = df[df['Order No'] == order_no]
        
        for _, row in order_df.iterrows():
            item_line = ET.SubElement(sales_order, "ITEMLINE", operation="Add")
            ET.SubElement(item_line, "KeyID")
            ET.SubElement(item_line, "ITEMNO").text = safe_str(row['Item Code'])
            ET.SubElement(item_line, "QUANTITY").text = safe_str(row['Quantity'])
            ET.SubElement(item_line, "ITEMUNIT").text = safe_str(row['UOM'])
            ET.SubElement(item_line, "UNITRATIO")
            for i in range(1, 11):
                ET.SubElement(item_line, f"ITEMRESERVED{i}")
            ET.SubElement(item_line, "ITEMOVDESC").text = safe_str(row['Item Name'])
            ET.SubElement(item_line, "UNITPRICE").text = safe_str(row['Unit Price'])
            ET.SubElement(item_line, "DISCPC")
            ET.SubElement(item_line, "TAXCODES")
            ET.SubElement(item_line, "GROUPSEQ")
            ET.SubElement(item_line, "QTYSHIPPED")

        ET.SubElement(sales_order, "SONO").text = safe_str(order_no)
        ET.SubElement(sales_order, "SODATE").text = safe_str(row['Posted Date'])
        ET.SubElement(sales_order, "TAX1ID").text = "T"
        ET.SubElement(sales_order, "TAX1CODE").text = "T"
        ET.SubElement(sales_order, "TAX2CODE")
        ET.SubElement(sales_order, "TAX1RATE").text = "11"
        ET.SubElement(sales_order, "TAX2RATE").text = "0"
        ET.SubElement(sales_order, "TAX1AMOUNT")
        ET.SubElement(sales_order, "TAX2AMOUNT").text = "0"
        ET.SubElement(sales_order, "RATE").text = "1"
        ET.SubElement(sales_order, "TAXINCLUSIVE").text = "0"
        ET.SubElement(sales_order, "CUSTOMERISTAXABLE").text = "1"
        ET.SubElement(sales_order, "CASHDISCOUNT").text = safe_str(row['Discount'])
        ET.SubElement(sales_order, "CASHDISCPC")
        ET.SubElement(sales_order, "FREIGHT")
        ET.SubElement(sales_order, "TERMSID")
        ET.SubElement(sales_order, "SHIPVIAID")
        ET.SubElement(sales_order, "FOB")
        ET.SubElement(sales_order, "ESTSHIPDATE")
        ET.SubElement(sales_order, "DESCRIPTION")
        
        ET.SubElement(sales_order, "SHIPTO1").text = safe_str(row['Address'])
        ET.SubElement(sales_order, "SHIPTO2").text = safe_str(row['District'])
        ET.SubElement(sales_order, "SHIPTO3")
        ET.SubElement(sales_order, "SHIPTO4")
        ET.SubElement(sales_order, "SHIPTO5")
        ET.SubElement(sales_order, "DP")
        ET.SubElement(sales_order, "DPACCOUNTID")
        ET.SubElement(sales_order, "DPUSED")
        ET.SubElement(sales_order, "CUSTOMERID").text = safe_str(row['Customer Code'])
        ET.SubElement(sales_order, "PONO")
        ET.SubElement(sales_order, "CURRENCYNAME").text = "IDR"

    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="    ")
    with open("output.xml", "w", encoding="utf-8") as f:
        f.write(xml_str)
    print("XML file created successfully.")

create_sales_order_xml(df)
