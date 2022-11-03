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


namespace KetCauThepNhat
{
    public class WblockBanVe
    {
        [CommandMethod("Exportdrawing")]
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
                int sohieubanve = 1 ;
                foreach (var itemId in entIds)
                {
                 
                    var entity = acTrans.GetObject(itemId, OpenMode.ForRead) as Entity;
                    if (entity != null)
                    {
                        var minPoint = entity.GeometricExtents.MinPoint;
                        var maxPoint = entity.GeometricExtents.MaxPoint;
                        //int n = 1;

                        using (DocumentLock acLckDocCur = acDoc.LockDocument())
                        {

                            #region tao vung chon du lieu copy
                            // Tao vung trong bang cach chon khung ban ve
                            PromptSelectionResult selectFraming;
                            selectFraming = acDocEd.SelectCrossingWindow(minPoint,
                                                                        maxPoint);
                            // If the prompt status is OK, objects were selected
                            if (selectFraming.Status != PromptStatus.OK)
                                return;
                            // Dat dia chi ID cho cac doi tuong duoc chon 
                            var entsFraming = selectFraming.Value;
                            var entIdsFraming = entsFraming.GetObjectIds();

                            acObjIdColl = new ObjectIdCollection();
                            foreach (var entIdsF in entIdsFraming)
                            {
                                var entity1 = acTrans.GetObject(entIdsF, OpenMode.ForRead) as Entity;
                                if (entity1 != null)
                                {

                                    // Add all the objects to copy to the new document

                                    acObjIdColl.Add(entIdsF);


                                }
                            }
                            #endregion
                        }
                        #region Tao file moi
                        // Change the file and path to match a drawing template on your workstation
                        string sLocalRoot = Application.GetSystemVariable("LOCALROOTPREFIX") as string;
                        string sTemplatePath = sLocalRoot + "Template\\mizonishi.dwt";

                        // Create a new drawing to copy the objects to
                        DocumentCollection acDocMgr = Application.DocumentManager;
                        Document acNewDoc = acDocMgr.Add(sTemplatePath);
                        Database acDbNewDoc = acNewDoc.Database;
                        
                        // Lock the new document
                        using (DocumentLock acLckDoc = acNewDoc.LockDocument())
                        {
                            // Start a transaction in the new database
                            using (Transaction acTrans1 = acDbNewDoc.TransactionManager.StartTransaction())
                            {
                                acDbNewDoc.Ltscale = 12;
                                // Open the Block table for read
                                BlockTable acBlkTblNewDoc;
                                acBlkTblNewDoc = acTrans1.GetObject(acDbNewDoc.BlockTableId,
                                                                    OpenMode.ForRead) as BlockTable;

                                // Open the Block table record Model space for read
                                BlockTableRecord acBlkTblRecNewDoc;
                                acBlkTblRecNewDoc = acTrans1.GetObject(acBlkTblNewDoc[BlockTableRecord.ModelSpace],
                                                                    OpenMode.ForRead) as BlockTableRecord;

                                // Clone the objects to the new database
                                IdMapping acIdMap = new IdMapping();
                                acCurDb.WblockCloneObjects(acObjIdColl, acBlkTblRecNewDoc.ObjectId, acIdMap,
                                                            DuplicateRecordCloning.Ignore, false);

                                // Save the copied objects to the database
                                acTrans1.Commit();
                            }

                        }
                        #region Luu Ban Ve

                        acDocMgr.CurrentDocument = acNewDoc;

                        //acDocMgr.MdiActiveDocument = acNewDoc;
                        string strDWGName = acNewDoc.Name;

                        // then provide a new name
                        var tenbanve = sohieubanve;

                        strDWGName = "C:\\SAVE\\MyDrawing" + tenbanve + ".dwg";
                        //
                        
                        // Save the active drawing
                        //Application.DocumentManager.CurrentDocument.Database.SaveAs(strDWGName, true, DwgVersion.Current,
                        //                      acNewDoc.Database.SecurityParameters);
                        Application.DocumentManager.CurrentDocument.Database.SaveAs(strDWGName, DwgVersion.Current);
                        Application.DocumentManager.CurrentDocument.CloseAndDiscard();

                        sohieubanve = sohieubanve + 1;


                        #endregion

                        // Unlock the document


                        // Set the new document current
                        //acDocMgr.MdiActiveDocument = acNewDoc;

                        acDocMgr.CurrentDocument = acDoc;
                        #endregion 


                    }

                }

                acTrans.Commit();
            }

        }
    }
}
