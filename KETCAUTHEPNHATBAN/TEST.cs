using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace KETCAUTHEPNHATBAN
{
    public class  TEST
    {
        [CommandMethod("TEST")]
        public static void Exportdrawing()
        {
            // Get the current document editor
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Database acCurDb = acDoc.Database;
            ObjectIdCollection acObjIdColl = new ObjectIdCollection();

            PromptSelectionOptions selectOption = new PromptSelectionOptions();
            selectOption.MessageForAdding = "\n枠図面を選択する";

            TypedValue[] types = new TypedValue[2];

            var typeOption1 = new TypedValue((int)DxfCode.Start, "LWPOLYLINE,POLYLINE");
            var typeOption2 = new TypedValue((int)DxfCode.LayerName, "Print");

            types.SetValue(typeOption1, 0);
            types.SetValue(typeOption2, 1);

            // Assign the filter criteria to a SelectionFilter object

            var filter = new SelectionFilter(types);

            // Request for objects to be selected in the drawing area

            var result = acDoc.Editor.GetSelection(selectOption, filter);

            if (result.Status != PromptStatus.OK) return;
            var ents = result.Value;
            var entIds = ents.GetObjectIds();


            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                var acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                            OpenMode.ForRead) as BlockTable;

                var acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;
                // Xac dinh toa do khung ban ve
                int sohieubanve = 1;
                foreach (var itemId in entIds)
                {

                    var entity = acTrans.GetObject(itemId, OpenMode.ForRead) as Entity;
                    if (entity != null)
                    {
                        var minPoint = entity.GeometricExtents.MinPoint;
                        //var maxPoint = entity.GeometricExtents.MaxPoint;
                        //int n = 1;
                        List<Point3d> pointsOfList = new List<Point3d>();
                        pointsOfList.Add(minPoint);


                    }

                }

                acTrans.Commit();
            }

        }
    }
}
