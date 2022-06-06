using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;

namespace LocalNetworkLogger
{
    public class PlaneVisiblityHandlerOnDrawing
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

        public void PlaneVisiblityHandler()
        {
            DrawingDocument oDrawingDoc = (DrawingDocument)InventorApp.ActiveDocument;
            Sheets oSheets = oDrawingDoc.Sheets;
            Sheet oSheet = oDrawingDoc.ActiveSheet;
            DrawingView oDrawView = InventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select Drawing view");

            if (oDrawView == null)
            {
                return;
            }

            AssemblyDocument oAssyDoc = oDrawView.ReferencedDocumentDescriptor.ReferencedDocument;
            AssemblyComponentDefinition assyCompDef = oAssyDoc.ComponentDefinition;
            WorkPlanes oWorkPlanes = assyCompDef.WorkPlanes;
            WorkPlane oWrkPlane = null;
            foreach(WorkPlane oWorkPlane in oWorkPlanes)
            {
                if(oWorkPlane.Name == "YZ Plane")
                {
                   oWorkPlane.AutoResize = true;
                    oWrkPlane = oWorkPlane;
                    break;
                }
            }
            oDrawView.SetIncludeStatus(oWrkPlane,true);
           
        }
    }
}
