using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Inventor;
using System.Diagnostics;

namespace LocalNetworkLogger
{
    internal class BOMHandler
    {
        private static Inventor.Application _inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");

        public static Inventor.Application InventorApp
        {
            get
            {
                if (_inventorApp == null)
                {
                    _inventorApp = GetInventorObj();
                }
                return _inventorApp;
            }
            set { _inventorApp = value; }
        }
        public List<string> itemNo = new List<string>();
        public List<int> itemQty = new List<int>();
        public List<string> partNo = new List<string>();
        public List<string> Desc = new List<string>();
        public void ExtractBOM()
        {
            AssemblyDocument assembly = (AssemblyDocument)InventorApp.ActiveDocument;
            AssemblyComponentDefinition assemCompDef = assembly.ComponentDefinition;
            BOM oBOM = assemCompDef.BOM;
            oBOM.StructuredViewFirstLevelOnly = false;
            oBOM.StructuredViewDelimiter = ".";
            oBOM.StructuredViewEnabled = true;
            BOMView oBOMView= null;
            foreach (BOMView x in oBOM.BOMViews)
            {
                if (x.Name == "Structured")
                {
                    oBOMView = x;
                }
            }
            //FileFormatEnum oFileFormat = FileFormatEnum.kMicrosoftExcelFormat;
            //oBOMView.Export(@"D:\TestBOM.xls", oFileFormat);
            GetBOMData(oBOMView.BOMRows);
            PrintBOM();
        }

       public List<BOMRowsEnumerator> BOMRowsEnum = new List<BOMRowsEnumerator>();
        public void GetBOMData(BOMRowsEnumerator oBOMRows)
        {            
            for (int i = 1; i <= oBOMRows.Count; i++)
            {
                BOMRow oRow = oBOMRows[i];
                if (oRow.ChildRows == null)
                {
                    GetRowData(oRow);
                }
                else
                {
                    GetRowData(oRow);
                    GetChildBOMData(oRow.ChildRows);
                }
            }

        }
        private void GetChildBOMData(BOMRowsEnumerator oBOMRows)
        {

            for (int i = 1; i <= oBOMRows.Count; i++)
            {
                BOMRow oRow = oBOMRows[i];
                if (oRow.ChildRows == null)
                {
                    GetRowData(oRow);
                }
                else
                {
                    GetRowData(oRow);
                    GetChildBOMData(oRow.ChildRows);
                }
            }

        }
        public void GetRowData(BOMRow xRow)
        {
            ComponentDefinition oCompDef = null;
            if (xRow.ItemQuantity == 0)
                return;

            oCompDef = xRow.ComponentDefinitions[1];
            
            PropertySet oPropSet = null;
            PropertySets oPropSets = oCompDef.Document.PropertySets;
            foreach (PropertySet xProp in oPropSets)
            {
                if (xProp.Name == "Design Tracking Properties")
                {
                    oPropSet = xProp;
                    if (oPropSet != null)
                        break;
                }
            }
            itemNo.Add(xRow.ItemNumber);
            //Debug.Print(itemNo.Last().ToString());
            itemQty.Add(xRow.ItemQuantity);
            bool foundPart = false;
            bool foundDesc = false;
            foreach (Property yProp in oPropSet)
            {
                if (yProp.Name == "Part Number")
                {
                    partNo.Add(yProp.Value);
                    foundPart = true;
                    //Debug.Print(partNo.Last().ToString());
                }
                else if (yProp.Name == "Description")
                {
                    Desc.Add(yProp.Value);
                    foundDesc = true;
                }
                if (foundPart && foundDesc)
                    break;
            }
        }

        private static Inventor.Application GetInventorObj()
        {
            Inventor.Application inventorApp = null;
            try
            {
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch { }
            return inventorApp;
        }
        public void PrintBOM()
        {
            for(int i =0;i<partNo.Count;i++)
            {
                string coll = itemNo[i].ToString()+"  "+itemQty[i].ToString()+"  "+partNo[i].ToString()+"  "+Desc[i].ToString();
                Debug.Print(coll);
            }
        }
    }

}
