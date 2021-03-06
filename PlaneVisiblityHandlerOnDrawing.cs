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
            ComponentOccurrences Occs = assyCompDef.Occurrences;
            ComponentOccurrencesEnumerator leaf = Occs.AllLeafOccurrences;

            foreach (ComponentOccurrence Occ in leaf)
            {
                dynamic occCompDef = null;
                if (Occ.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                    occCompDef = (PartComponentDefinition)Occ.Definition;
                else if (Occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    occCompDef = (AssemblyComponentDefinition)Occ.Definition;


                foreach (WorkPlane occPlane in occCompDef.WorkPlanes)
                {
                    object proxyPlaneObj = null;
                    Occ.CreateGeometryProxy(occPlane, out proxyPlaneObj);
                    WorkPlaneProxy workPlaneProxy = (WorkPlaneProxy)proxyPlaneObj;
                    if(IncludePlaneObject(workPlaneProxy,oDrawView))
                    {
                        GetDrawingLinePosition(workPlaneProxy, oSheet, oDrawView);
                    }
                }
            }
        }

        public void GetDrawingLinePosition(dynamic oWorkPlane, Sheet oSheet, DrawingView oDrawView)
        {
            double viewHeight = Math.Round(oDrawView.Height,0);
            double viewWidth = Math.Round(oDrawView.Width,0);
            double viewTop =  Math.Round(oDrawView.Top, 0);
            double viewBottom = Math.Round(viewTop - viewHeight,0);
            double viewLeft =  Math.Round(oDrawView.Left, 0);
            Point2d drawPosition = oDrawView.Position;
            foreach(Centerline centerline in oSheet.Centerlines)
            {
                if(centerline.CenterlineType == CenterlineTypeEnum.kWorkFeatureCenterlineType)
                {
                    if(centerline.ModelWorkFeature.Name == oWorkPlane.Name)
                    {
                        if(Math.Round(centerline.StartPoint.X,0) >= viewLeft && Math.Round(centerline.StartPoint.X, 0) <= viewLeft+viewWidth && Math.Round(centerline.StartPoint.Y, 0) <= viewTop && Math.Round(centerline.StartPoint.Y, 0) >= viewBottom)
                        {
                            oDrawView.SetVisibility(oWorkPlane, true);                            
                        }
                        else
                        {
                            oDrawView.SetIncludeStatus(oWorkPlane, false);
                        }
                    }
                }
            }

        }


        public static bool IncludePlaneObject(dynamic wPlane, DrawingView drawingView)
        { 
            try
            {
                drawingView.SetIncludeStatus(wPlane, true);

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
    }
}
