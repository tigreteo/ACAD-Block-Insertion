using System;
using System.IO;
using System.Text;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace BlockInsertionV3
{
    //code can occassionally break on particular arcs, hard to recreate error!!!!!!!!!!!!
    //could not re-create error on same drawing
    class FetchInsertArea
    {
        public static Region pullSideData(string specPath)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            Region returnRegion;

            //error log
            string errLog = @"C:\temp\Error Log.csv";
            StringBuilder attOut = new StringBuilder();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //open block table for read
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                //load sideDB
                Database sideDB = new Database(false, true);
                //load sideDB
                try { sideDB.ReadDwgFile(specPath, FileShare.ReadWrite, false, ""); }
                catch (System.Exception e)
                {
                    //alert user to error
                    ed.WriteMessage("\nUnable to read: " + specPath);
                    ////log error
                    attOut.AppendLine("Unable to read: " + specPath);
                    File.AppendAllText(errLog, attOut.ToString());
                    return null;
                }
                //insert into current DB to get data
                db.Insert(Path.GetFileNameWithoutExtension(specPath), sideDB, true);
                ObjectId blkRecId = bt[Path.GetFileNameWithoutExtension(specPath)];

                //get tables from sideDB
                BlockTableRecord sideBTR = tr.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                //convert selection to dboCollection exploding any polylines
                ObjectIdCollection deletePls = new ObjectIdCollection();
                DBObjectCollection parts = new DBObjectCollection();

                //add all viable parts to collection to make regions
                foreach (ObjectId objId in sideBTR)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    //get pieces of polyline
                    if (ent.ObjectId.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE")
                    {
                        //mark for delete from collection
                        deletePls.Add(objId);
                        Polyline ply = tr.GetObject(objId, OpenMode.ForRead) as Polyline;
                        DBObjectCollection plyprts = new DBObjectCollection();
                        ply.Explode(plyprts);

                        foreach (DBObject obj in plyprts)
                        { parts.Add(obj); }
                    }
                    //break up circle into arcs
                    if (ent.ObjectId.ObjectClass.DxfName.ToUpper() == "CIRCLE")
                    {
                        //mark for delete from collection
                        deletePls.Add(objId);

                        Circle cir = tr.GetObject(objId, OpenMode.ForRead) as Circle;
                        //create two arcs as half of each circle
                        Arc leftHemi = new Arc(cir.Center, cir.Radius, Math.PI / 2, (3 * Math.PI) / 2);
                        Arc rightHemi = new Arc(cir.Center, cir.Radius, (3 * Math.PI) / 2, Math.PI / 2);
                        parts.Add(leftHemi);
                        parts.Add(rightHemi);
                    }

                    DBObject dbo = tr.GetObject(objId, OpenMode.ForRead) as DBObject;
                    parts.Add(dbo);
                }
                //remove polylines from original collection
                foreach (ObjectId objId in deletePls)
                {
                    DBObject removeMe = tr.GetObject(objId, OpenMode.ForRead) as DBObject;
                    parts.Remove(removeMe);
                }

                //Test each object for an intersection of the other parts
                //at any intersections pieces will be split into new segments
                int i = -1;
                while (++i < parts.Count)
                {
                    Curve segi = parts[i] as Curve;
                    if (segi == null)
                        continue;

                    int j = -1;
                    while (++j < parts.Count)
                    {
                        if (i == j)
                            continue;

                        Curve segj = parts[j] as Curve;
                        if (segj == null)
                            continue;

                        //test segs against each other
                        Point3dCollection intersecPnts = new Point3dCollection();
                        segi.IntersectWith(segj, Intersect.OnBothOperands, intersecPnts, IntPtr.Zero, IntPtr.Zero);

                        //NEED TO ACCOUNT FOR ANY ARCS
                        //sample code had FuzzyEquals
                        for (int k = intersecPnts.Count - 1; k >= 0; k--)
                        {
                            Point3d pt = intersecPnts[k];
                            if (Equals(pt, segi.StartPoint) || Equals(pt, segi.EndPoint))
                                intersecPnts.RemoveAt(k);
                        }

                        if (intersecPnts.Count > 0)
                        {
                            var splCurves = segi.GetSplitCurves(intersecPnts);
                            //dont add curves split at start or end
                            if (splCurves.Count > 1)
                            {
                                foreach (DBObject dbo in splCurves)
                                {
                                    parts.Add(dbo);
                                }
                                parts[i] = null;
                                break;//exit while j, next i
                            }
                        }
                    }
                }

                //create regions from segments
                var regions = Region.CreateFromCurves(parts);

                //using the second largest region found bc of the nature of how we typically draw  forms.
                //API tends to leave outer lines as a region, even though it clearly intersects with other regions

                //loop through regions to find second largest
                Region largest = new Region(); Region secLarge = new Region();
                foreach (Region dbo in regions)
                {
                    if (dbo.Area >= largest.Area)
                    {
                        secLarge = largest;
                        largest = dbo;
                    }
                    else if (dbo.Area < largest.Area && dbo.Area > secLarge.Area)
                    { secLarge = dbo; }
                }

                returnRegion = secLarge;
                //get bounds of the region found
                Point3d origin = Point3d.Origin;
                Vector3d xAxis = Vector3d.XAxis;
                Vector3d yAxis = Vector3d.YAxis;
                Extents2d dropAreaExt = secLarge.AreaProperties(ref origin, ref xAxis, ref yAxis).Extents;
                Point2d centroid = secLarge.AreaProperties(ref origin, ref xAxis, ref yAxis).Centroid;                

                //standard insert point is 0,0,0
                //to find this point it may require a sysVariable check

                tr.Commit();
            }
            //return this data
            //return region and use data accordingly
            return returnRegion;
        }
    }
}
