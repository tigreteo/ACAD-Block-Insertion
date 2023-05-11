using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Linq;

namespace BlockInsertionV3
{
    class EntityFromPoint
    {
        /// <summary>
        /// Select entities by iterating over the model space.
        /// Can add filter for specific type of Entity
        /// </summary>
        public static ObjectId[] SelectCrossingDatabase(Point3d pnt, string entType = "NA")
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            try
            {
                ObjectIdCollection idCol = new ObjectIdCollection();

                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        Point3d nearPnt = new Point3d();
                        PointOnCurve3d clstPnt = new PointOnCurve3d();

                        if (id.ObjectClass.DxfName.ToUpper() == "POLYLINE")
                        {
                            Polyline pLine = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                            nearPnt = pLine.GetClosestPointTo(pnt, true);

                            if (pnt.DistanceTo(nearPnt) < 0.01)
                            {
                                idCol.Add(id);
                            }
                        }

                        DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                        Line dbLine = obj as Line;
                        if (dbLine != null)
                        {
                            LineSegment3d lineSegment = new LineSegment3d(dbLine.StartPoint, dbLine.EndPoint);
                            PointOnCurve3d q = lineSegment.GetClosestPointTo(pnt);
                            if (pnt.DistanceTo(q.Point) < 0.01)
                            {
                                idCol.Add(id);
                            }
                        }
                    }

                    if (idCol.Count > 0)
                    {
                        ObjectId[] ids = new ObjectId[idCol.Count];
                        int counter = 0;
                        foreach (ObjectId id in idCol)
                        {
                            ids[counter] = id;
                            counter++;
                        }

                        if (entType != "NA")
                        {
                            //filter
                            ids = onlyType(ids, entType);
                        }
                        tr.Commit();
                        return ids;
                    }
                    else
                    {
                        tr.Commit();
                        return null;
                    }
                }
            }
            catch (System.Exception e)
            {
                ed.WriteMessage("\nException {0}.", e);
                return null;
            }
        }

        /// <summary>
        /// Create window around point to find parts in area of .01^2
        /// </summary>
        /// <param name="pnt"></param>
        /// <param name="entType"></param>
        /// <returns></returns>
        public static ObjectId[] SelectCrossingWindow(Point3d pnt, string entType = "NA")
        {
            Point3d p1 = new Point3d(pnt.X - 0.01, pnt.Y - 0.01, 0);
            Point3d p2 = new Point3d(pnt.X + 0.01, pnt.Y + 0.01, 0);

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptSelectionResult res = ed.SelectCrossingWindow(p1, p2);
            SelectionSet sel = res.Value;


            ObjectId[] ids = new ObjectId[sel.Count];
            int counter = 0;
            foreach (ObjectId id in sel.GetObjectIds())
            {
                ids[counter] = id;
                counter++;
            }

            if (entType != "NA")
            {
                //filter
                ids = onlyType(ids, entType);
            }

            return ids;
        }

        private static ObjectId[] onlyType(ObjectId[] ids, string entType)
        {
            return
                ids.Where(id => { return id.ObjectClass.DxfName.ToUpper() == entType; }).ToArray();
        }
    }
}
