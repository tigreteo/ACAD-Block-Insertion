using System;
using System.Text;
using System.IO;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;

//Use to insert/maintain spec sheets in autocad
namespace BlockInsertionV3
{
    public class Class1
    {
        //insertSpec command
        [CommandMethod("InsertSpec")]
        public void Insertthis()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //request the selection
                SelectionSet selSet;
                PromptSelectionResult psr = ed.GetSelection();
                if (psr.Status != PromptStatus.OK)
                { return; }
                else
                { selSet = psr.Value; }

                //get extents of selSet
                Extents3d selectDims = TransactionExtensions.GetExtents(tr, selSet);

                //get spec filePath
                string specPath = getSpec(doc, ed);
                //get spec name
                string spec = Path.GetFileNameWithoutExtension(specPath);

                #region special Conditions
                //if spec == fabric layout
                //get dimensions text from selection and use the biggest measurement to fillout the block
                Fraction yardage = new Fraction();
                if (spec.ToUpper() == "FABRIC LAYOUT")
                {
                    //get dimensions
                    double dimYards = dimsfromSelect(selSet, tr, ed);
                    if (dimYards != 0)
                    { yardage = inchestoYards.makeFraction(dimYards); }
                }

                //if spec == sew spec
                //check if there are sew panels in selection and handle that differently
                //reScale, align, move sew panels
                if (spec.ToUpper() == "SEWING SPEC")
                {

                }
                #endregion

                //get relevant spec dimensions
                Extents3d dropArea = getSpecDim(specPath);

                //get scale
                double scale = calcScale(selectDims, dropArea);

                //calculate insertpoint
                Point3d insertPoint = getInsertPoint(scale, dropArea, selectDims);

                //start insertion
                BlockInserter(specPath, ed, insertPoint, scale, yardage, selectDims);

                //commit changes
                tr.Commit(); 
            }
        }

        //command is used to better accomadate a spec(blockRef) around its drawing
        //Request selection set from user
        //Request diagonal of target area
        //code derives scale difference between two different areas (rounds scale to real number)
        //using chosen area, find the related block (possibly let user select?)
        //default mode moves selection to scaled spec (optionally can move scaled spec to selection)
        //TODO-Needs to accomadate dynamic blocks
        [CommandMethod("UpdateScale")]
        public void updateSpecScale()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            bool mode = true;

            //start Transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                PromptPointOptions ppo = new PromptPointOptions("");
                ppo.Message = "Choose start of diagonal";
                //get keywords set up
                ppo.Keywords.Add("StartPoint");
                ppo.Keywords.Add("EndPoint");
                ppo.Keywords.Add("Mode");
                ppo.Keywords.Add("Selection");

                PromptSelectionOptions pso = new PromptSelectionOptions();
                //get keywords set up
                pso.Keywords.Add("StartPoint");
                pso.Keywords.Add("EndPoint");
                pso.Keywords.Add("Mode");
                pso.Keywords.Add("Selection");

                //set prompts to include keywords
                string kws = pso.Keywords.GetDisplayString(true);

                //implement a callback for when keywords are entered
                pso.KeywordInput +=
                    delegate (object send, SelectionTextInputEventArgs e)
                    {
                        ed.WriteMessage("\nKeyword entered: {0}", e.Input);
                    };

                //in original selection have option to change mode, but default mode is selection moved to spec
                //get selection set from user
                SelectionSet selSet;
                PromptSelectionResult psr = ed.GetSelection();
                if (psr.Status != PromptStatus.OK)
                { return; }
                else
                { selSet = psr.Value; }

                //get a window of the selection set
                Extents3d selectExtns = tr.GetExtents(selSet);

                //get two points from user                
                Point3d strtPoint = new Point3d();
                Point3d endPoint = new Point3d();
                PromptPointResult ppr = ed.GetPoint("Choose start of diagonal");
                if (ppr.Status != PromptStatus.OK)
                { return; }
                else
                    strtPoint = ppr.Value;

                ppr = ed.GetPoint("Choose end of diagonal");
                if (ppr.Status != PromptStatus.OK)
                { return; }
                endPoint = ppr.Value;
                //Diagonal line represents target area
                Extents3d diagonal = coordinates(strtPoint, endPoint);

                //find spec using startpoint
                ObjectId[] ids = EntityFromPoint.SelectCrossingWindow(strtPoint);
                BlockReference blkRef = tr.GetObject(ids[0], OpenMode.ForWrite) as BlockReference;

                //get scale from blkref use to work back to size at scale 1, to then apply to new calc scale
                double origScale = blkRef.ScaleFactors.X;
                Point3d blkCorner = blkRef.Position;

                //calc the center of the diagonal provided by user
                Point3d specCenter = getCentroid(diagonal.MinPoint, diagonal.MaxPoint);
                //calc the center of the selected area
                Point3d selectionCenter = getCentroid(selectExtns.MinPoint, selectExtns.MaxPoint);

                //adjust SpecCenter to be relative to insertPoint of blockRef
                specCenter = new Point3d(
                    specCenter.X - blkCorner.X,
                    specCenter.Y - blkCorner.Y,
                    specCenter.Z - blkCorner.Z);

                //adjust spec target center to when scale is 1 for algo
                specCenter = new Point3d(
                    specCenter.X / origScale,
                    specCenter.Y / origScale,
                    specCenter.Z / origScale);


                //calculate scale as though block is at scale 1
                double scale = calcScale(selectExtns, diagonal);

                //transform blockRef
                blkRef.TransformBy(Matrix3d.Scaling(scale / origScale, blkCorner));
                //adjust spec target center to new scale
                specCenter = new Point3d(
                    specCenter.X * scale,
                    specCenter.Y * scale,
                    specCenter.Z * scale);

                //add spec center to the corner of the block
                specCenter = new Point3d(
                    specCenter.X + blkCorner.X,
                    specCenter.Y + blkCorner.Y,
                    specCenter.Z + blkCorner.Z);

                if (mode)
                    //move selection set to spec
                    foreach (ObjectId id in selSet.GetObjectIds())
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        ent.TransformBy(Matrix3d.Displacement(selectionCenter.GetVectorTo(specCenter)));
                    }
                else
                { blkRef.TransformBy(Matrix3d.Displacement(specCenter.GetVectorTo(selectionCenter))); }

                tr.Commit();
            }

        }

        //get which spec sheet is being used
        //1 assume spec on current folderpath
        //2 request dialoge from user with pre-suggested specs
        //3 other option in dialoge opens filepath dialoge
        //Need to return spec filepath
        private string getSpec(Document doc, Editor ed)
        {
            string specName = "default";
            //check to see if the file has not been named yet
            if (System.Convert.ToInt16(Application.GetSystemVariable("DWGTITLED")) != 0)
            {
                string fileName = doc.Name;
                string[] nameParts = fileName.Split('\\');
                //back up to the specname in typical naming scheme
                specName = nameParts[nameParts.Length - 2];
            }

            //if not named yet, or specname isn't known we need to pick one
            switch (specName.ToUpper())
            {
                //validate that it is a spec we expect
                //instead of comparing to this list, might want to compare to the spec sheet folder
                case "FABRIC":
                case "PATTERN":
                case "CARDBOARD":
                case "CARD BOARD":
                case "POLY":
                    break;
                default:
                    //request polySpec
                    PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                    pKeyOpts.Message = "\nEnter spec type ";
                    pKeyOpts.Keywords.Add("Pattern");
                    pKeyOpts.Keywords.Add("Cardboard");
                    pKeyOpts.Keywords.Add("Sewing");
                    pKeyOpts.Keywords.Add("SEWPanel");
                    pKeyOpts.Keywords.Add("Poly");
                    pKeyOpts.Keywords.Add("Frame");
                    pKeyOpts.Keywords.Add("Other");

                    PromptResult res = ed.GetKeywords(pKeyOpts);
                    if (res.Status == PromptStatus.OK)
                        specName = res.StringResult.ToUpper();
                    break;
            }

            //change this should the folder move or the server change namesVVVVVVVVVVVV
            string folder = @"Y:\Product Development\Forms\Spec Sheets\";
            string specSheet = "";
            switch (specName.ToUpper())
            {
                case "FABRIC":
                case "PATTERN":
                    specSheet = "Fabric Layout.dwg";
                    break;
                case "CARDBOARD":
                case "CARD BOARD":
                    specSheet = "CardboardForm.dwg";
                    break;
                case "SEWING":
                    specSheet = "Sewing Spec.dwg";
                    break;
                case "SEWPANEL":
                    specSheet = "Sewing Spec Panel Dynamic.dwg";
                    break;
                case "POLY":
                    specSheet = "Poly Specification.dwg";
                    break;
                case "FRAME":
                    return @"Y:\Engineering\Drawings\Blocks\Standard Forms\SF_CNC_BORDER_V2.dwg";
                default:
                    //call method to choose a file from the folder, or some such
                    specSheet = pickSpec(ed, doc);
                    return specSheet;
            }
            specSheet = folder + specSheet;
            return specSheet;
        }

        //continue requesting a selection until only ONE dimension is found
        private static SelectionSet pickOne(SelectionSet acSet, Transaction tr, Editor ed, SelectionFilter sFilter, bool solved)
        {
            if (solved == false)
            {
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nChoose the (1) dimension for the layout:";
                PromptSelectionResult psr = ed.GetSelection(pso, sFilter);
                if (psr.Status == PromptStatus.OK)
                {
                    acSet = psr.Value;
                    if (acSet.Count != 1)
                    { acSet = pickOne(acSet, tr, ed, sFilter, solved); }

                    return acSet;
                }
                return null;
            }
            else
                return acSet;
        }

        //sort through a selection to get dimensions from a dimension if it exists
        private double dimsfromSelect(SelectionSet acSet, Transaction tr, Editor ed)
        {
            //make a filter that only allows selections of dimensions
            TypedValue[] acValue = new TypedValue[1];
            acValue.SetValue(new TypedValue((int)DxfCode.Start, "DIMENSION"), 0);
            SelectionFilter sFilter = new SelectionFilter(acValue);

            double dimensionYardage = 0;
            ObjectIdCollection dims = new ObjectIdCollection();
            foreach (ObjectId id in acSet.GetObjectIds())
            {
                if (id.ObjectClass.DxfName.ToUpper() == "DIMENSION")
                { dims.Add(id); }
            }
            //IF there is more than one dimension. ask the user to pick one
            if (dims.Count > 0)
            {
                acSet = pickOne(acSet, tr, ed, sFilter, false);
                if (acSet == null)
                { return 0; }

                foreach (ObjectId objId in acSet.GetObjectIds())
                {
                    Dimension dim = tr.GetObject(objId, OpenMode.ForRead) as Dimension;
                    if (dim.Measurement > dimensionYardage)
                    { dimensionYardage = dim.Measurement; }
                }
            }

            return dimensionYardage;
        }

        //might need to adjust scale to round to .5 instead of whole numbers
        //can send a coeffecient to account for dynamic blocks that can have their load area linestretched
        public double calcScale(Extents3d selectDims, Extents3d dropArea, double xCoef = 1, double yCoef = 1)
        {
            double distX = Math.Abs(selectDims.MaxPoint.X - selectDims.MinPoint.X);
            double distY = Math.Abs(selectDims.MaxPoint.Y - selectDims.MinPoint.Y);
            //add on 10% for some margin
            distX = distX * 1.1;
            distY = distY * 1.1;

            //get dimensions of insert area of spec sheet
            double frameDistX = Math.Abs(dropArea.MaxPoint.X - dropArea.MinPoint.X);
            double frameDistY = Math.Abs(dropArea.MaxPoint.Y - dropArea.MinPoint.Y);

            //neccessary ratios can be offset by dynamic stretch axes
            double scaleX = distX / (frameDistX * xCoef);
            double scaleY = distY / (frameDistY * yCoef);

            //whichever is higher is the scale to work from
            double scale;
            if (scaleX > scaleY)
                scale = scaleX;
            else
                scale = scaleY;

            //round up to 1/2 unit       
            scale = (Math.Ceiling(2 * scale)) / 2;

            return scale;
        }

        //Needs to assume insertPoint is 0,0,0 for now, may be able to use sysVariable to find it
        private static Point3d getInsertPoint(double scale, Extents3d dropArea, Extents3d selectDims)
        {
            //get centroid of drop area
            Point3d targetCenter = getCentroid(dropArea.MinPoint, dropArea.MaxPoint);

            //get centroid of selection
            Point3d selectionCenter = getCentroid(selectDims.MinPoint, selectDims.MaxPoint);

            //using the distance of the center of the specsheet insert area from its ref point (0,0,0)
            //multiply said distance by the scale factor
            //use this new distance to offset from our selection's center for the insert point
            double x = selectionCenter.X - (targetCenter.X * scale);
            double y = selectionCenter.Y - (targetCenter.Y * scale);

            return new Point3d(x, y, 0);
        }

        //find center of an area defined by min and max points
        private static Point3d getCentroid(Point3d minPt, Point3d maxPt)
        {
            //formula
            //total distance, divided by two, added to original point
            double distX = (Math.Abs(maxPt.X - minPt.X) / 2) + minPt.X;
            double distY = (Math.Abs(maxPt.Y - minPt.Y) / 2) + minPt.Y;
            //could do the same for Z, but probaly not necissary

            return new Point3d(distX, distY, 0);
        }

        //gets the second to last fold in the path to assume styleID number
        private static string styleIdFromPath(Document doc)
        {
            try
            {
                string styleID = "";
                string pathName = Path.GetDirectoryName(doc.Name);
                string[] styleIDparts = pathName.Split('\\');
                styleID = styleIDparts[styleIDparts.Length - 2];
                return styleID;
            }
            catch
            { return null; }
        }

        //get style ID from fileName, can be changed to return group# or description#
        private static string styleIdFromName(string docName)
        {
            //find expected styleID and group number from filename
            string fileName = Path.GetFileNameWithoutExtension(docName);
            StringBuilder group = new StringBuilder();
            StringBuilder style = new StringBuilder();
            bool secondPart = false;
            bool firstPart = false;
            char[] groupList = { ' ', '-', 'A', 'C' };
            char[] styleList = { ' ', 'S', 'N' };//might need to change this to be more general
                                                 //it assumes and S as in SAM1, or N as in NESTERWOOD

            //loop through file name to generate parts of name
            //first to last
            foreach (char c in fileName)
            {
                //if number is false it isnt complete
                if (!secondPart)
                {
                    if (!firstPart)
                    {
                        //if part isnt one of the separaters add it to end of part
                        if (Array.Exists(groupList, element => element == c))
                        {
                            //if it is a separater then the first part is complete(true)
                            firstPart = true;
                        }
                        else
                            group.Append(c);
                    }

                    else {
                        //if par isnt one of the expected closers then add it to the end of the part
                        if (Array.Exists(styleList, element => element == c))
                        {
                            //if it is an expected closer then set second part to true
                            secondPart = true;
                        }
                        else
                            style.Append(c);
                    }
                }
            }
            string returnName = group + "-" + style;
            return returnName;
        }

        //acquire the extents of the frame and convert them from UCS to DCS, in case of view rotation
        static public Extents3d coordinates(Point3d firstInput, Point3d secondInput)
        {
            double minX;
            double minY;
            double maxX;
            double maxY;
            double minZ;
            double maxZ;

            //sort through the values to be sure that the correct first and second are assigned
            if (firstInput.X < secondInput.X)
            { minX = firstInput.X; maxX = secondInput.X; }
            else
            { maxX = firstInput.X; minX = secondInput.X; }

            if (firstInput.Y < secondInput.Y)
            { minY = firstInput.Y; maxY = secondInput.Y; }
            else
            { maxY = firstInput.Y; minY = secondInput.Y; }

            if (firstInput.Z < secondInput.Z)
            { minZ = firstInput.Z; maxZ = secondInput.Z; }
            else
            { maxZ = firstInput.Z; minZ = secondInput.Z; }


            Point3d first = new Point3d(minX, minY, minZ);
            Point3d second = new Point3d(maxX, maxY, maxZ);
            Extents3d window = new Extents3d(first, second);

            return window;
        }

        //use the windows dialog to have user choose the specific file
        public static string pickSpec(Editor ed, Document doc)
        {
            string spec = "";

            OpenFileDialog ofd = new OpenFileDialog(
                "Select a spec sheet to import", null,
                "dwg",
                "InsertSpec", OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);

            
            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
                spec = ofd.Filename;

            return spec;
        }

        //if data is known already(stored in SQL or registry)
        //else pull data from side DB (unless its already loaded)-update stored data
        //if derived data is null (getting data failed) use standard dimensions
        private static Extents3d getSpecDim(string specPath)
        {
            //currently looking for data in Registry
            //can be stored elsewhere, suggest SQL entries

            //registry key will be found in AutoCad folder using spec as name
            string regKey = @"Software\Autodesk\AutoCAD\" + Path.GetFileNameWithoutExtension(specPath);
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(regKey, true);

            bool updatePath = false;
            string minX = ".25";
            string maxX = "8.5";
            string minY = "0";
            string maxY = "4";
            //check if path exists
            if (key == null)
            {
                //create a key
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(regKey);
                //create a path to ref in the future as key value
                updatePath = true;
            }
            else
            {
                //see if date modified is as new or newer than that of the spec sheet file
                string dateModified = key.GetValue("Date").ToString();
                //find date of last modified
                FileInfo fi = new FileInfo(specPath);

                if(dateModified != null &&
                    Convert.ToDateTime(dateModified) > fi.LastWriteTime)
                {
                    //get region data and return
                    minX = key.GetValue("MinX").ToString();
                    maxX = key.GetValue("MaxX").ToString();
                    minY = key.GetValue("MinY").ToString();
                    maxY = key.GetValue("MaxY").ToString();
                }
                else
                {
                    updatePath = true;   
                }                
            }


            //update/create key data entries
            if (updatePath)
            {
                Region targetArea;
                //run command to get region data, put data into registry and return it
                try { targetArea = FetchInsertArea.pullSideData(specPath); }
                catch (System.Exception)
                {
                    //change to return default
                    return new Extents3d(
                        new Point3d(.25, 0,0),
                        new Point3d(8.5,4.5,0));
                }

                //get bounds of the region found
                Point3d origin = Point3d.Origin;
                Vector3d xAxis = Vector3d.XAxis;
                Vector3d yAxis = Vector3d.YAxis;
                Extents2d dropAreaExt = targetArea.AreaProperties(ref origin, ref xAxis, ref yAxis).Extents;

                key.SetValue("MinX", dropAreaExt.MinPoint.X);
                key.SetValue("MaxX", dropAreaExt.MaxPoint.X);
                key.SetValue("MinY", dropAreaExt.MinPoint.Y);
                key.SetValue("MaxY", dropAreaExt.MaxPoint.Y);
                key.SetValue("Date", DateTime.Now);


                Extents3d dropAreaQuick = new Extents3d(
                new Point3d(dropAreaExt.MinPoint.X, dropAreaExt.MinPoint.Y, 0),
                new Point3d(dropAreaExt.MaxPoint.X, dropAreaExt.MaxPoint.Y, 0));

                return dropAreaQuick;
            }

            Extents3d dropArea = new Extents3d(
                new Point3d(Convert.ToDouble(minX), Convert.ToDouble(minY), 0),
                new Point3d(Convert.ToDouble(maxX), Convert.ToDouble(maxY), 0));
            return dropArea;
        }

        //get relevant data to specific spec
        //Any dimensions ie LxW
        //Change logs
        //Special Notes
        //Style ID description
        //insert spec
        //center of drop area over(w/ repsect to insertpoint) center of selection
        //scale spec
        //scale using center drop area
        //if dynamic, stretch appropriate axis to fit selection + buffer & re-center
        //fill out data in spec

        //Gathers further relavant data and inserts the spec into the space
        public static void BlockInserter(string specPath, Editor ed, Point3d insertPoint, double scale, Fraction yardage, Extents3d selectDims,
            double xStretch = 0, double yStretch = 0)
        {
            string specName = Path.GetFileNameWithoutExtension(specPath);
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            //error log
            string errLog = @"C:\temp\Error Log.csv";
            StringBuilder attOut = new StringBuilder();

            //dimensions of selection
            double xDist = Math.Abs(selectDims.MaxPoint.X - selectDims.MinPoint.X);
            double yDist = Math.Abs(selectDims.MaxPoint.Y - selectDims.MinPoint.Y);

            //get styleID
            //have to decide between two different approaches
            //string styleId = styleIdFromName(specPath);
            string styleId = styleIdFromPath(Application.DocumentManager.MdiActiveDocument);

            //try to get the name for the styleID
            string idName = "";
            try
            { idName = ID_Reader.interpreter(styleId); }
            catch (System.Exception e)
            { ed.WriteMessage("\nUnknown StyleID name"); }

            //use the yardage struct to create an inches string
            StringBuilder sbInches = new StringBuilder();
            if (yardage.whole != 0)
            { sbInches.Append(yardage.whole.ToString()); }
            if (yardage.denom != 0 && yardage.num != 0)
            { sbInches.Append(" " + yardage.num.ToString() + "/" + yardage.denom.ToString()); }


            using (Transaction trCurrent = dbCurrent.TransactionManager.StartTransaction())
            {
                //open block table for read
                BlockTable btCurrent = trCurrent.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead) as BlockTable;

                //check if spec is already loaded into drawing
                ObjectId blkRecId = ObjectId.Null;
                if (!btCurrent.Has(specName))
                {
                    //open db to other file
                    Database db = new Database(false, true);
                    try
                    { db.ReadDwgFile(specPath, System.IO.FileShare.Read, false, ""); }
                    catch (System.Exception)
                    {
                        //alert user to error
                        ed.WriteMessage("\nUnable to read: " + specPath);
                        //log error
                        attOut.AppendLine("Unable to read: " + specPath);
                        File.AppendAllText(errLog, attOut.ToString());
                        return;
                    }
                    dbCurrent.Insert(specName, db, true);
                    blkRecId = btCurrent[specName];
                }
                else
                { blkRecId = btCurrent[specName]; }

                //now insert block into current space
                if (blkRecId != ObjectId.Null)
                {
                    //create btr for the inserted block
                    BlockTableRecord btrInsert = trCurrent.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                    using (BlockReference blkRef = new BlockReference(insertPoint, blkRecId))
                    {
                        BlockTableRecord btrCurrent = trCurrent.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        //scale the frame using insert point and scalefactor
                        blkRef.TransformBy(Matrix3d.Scaling(scale, insertPoint));

                        //add the frame to the btr
                        btrCurrent.AppendEntity(blkRef);
                        trCurrent.AddNewlyCreatedDBObject(blkRef, true);

                        // Verify block table record has attribute definitions associated with it
                        if (btrInsert.HasAttributeDefinitions)
                        {
                            // Add attributes from the block table record
                            foreach (ObjectId objID in btrInsert)
                            {
                                DBObject dbObj = trCurrent.GetObject(objID, OpenMode.ForRead) as DBObject;
                                if (dbObj is AttributeDefinition)
                                {
                                    AttributeDefinition acAtt = dbObj as AttributeDefinition;
                                    if (!acAtt.Constant)
                                    {
                                        using (AttributeReference acAttRef = new AttributeReference())
                                        {
                                            acAttRef.SetAttributeFromBlock(acAtt, blkRef.BlockTransform);
                                            acAttRef.Position = acAtt.Position.TransformBy(blkRef.BlockTransform);

                                            acAttRef.TextString = acAtt.TextString;

                                            blkRef.AttributeCollection.AppendAttribute(acAttRef);
                                            trCurrent.AddNewlyCreatedDBObject(acAttRef, true);
                                        }
                                    }
                                }
                            }
                            // Write new data into the block
                            //
                            AttributeCollection attCol = blkRef.AttributeCollection;
                            foreach (ObjectId objID in attCol)
                            {
                                DBObject dbObj = trCurrent.GetObject(objID, OpenMode.ForRead) as DBObject;
                                AttributeReference acAttRef = dbObj as AttributeReference;
                                //initials need to be in a specific file location or registry loc
                                if (acAttRef.Tag.Contains("DATE") && !acAttRef.Tag.Contains("APPROV"))
                                { acAttRef.TextString = DateTime.Now.ToShortDateString(); }
                                else if (acAttRef.Tag.Contains("ITEMNUM"))
                                { acAttRef.TextString = styleId; }
                                else if (acAttRef.Tag.Contains("ITEMDESC"))
                                { acAttRef.TextString = idName; }
                                else if (acAttRef.Tag.Contains("SCALE"))
                                { acAttRef.TextString = blkRef.ScaleFactors.X.ToString(); }
                                else if (acAttRef.Tag.Contains("YARDS2"))
                                {
                                    if (yardage.yds != 0)
                                    { acAttRef.TextString = yardage.yds.ToString(); }
                                }
                                else if (acAttRef.Tag.Contains("INCHES2"))
                                { acAttRef.TextString = sbInches.ToString(); }
                            }

                            //stretch parts if they are dynamic
                            DynamicBlockReferencePropertyCollection dynPropCol = blkRef.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty dynProp in dynPropCol)
                            {
                                if (dynProp.PropertyName.ToUpper() == "DISTANCE_DRAWAREA_HEIGHT")//stretch to be Height of selection + buffer %
                                {
                                    dynProp.Value = Convert.ToDouble(yDist * 1.1);
                                    //move to account for stretch change
                                    blkRef.TransformBy(Matrix3d.Displacement(new Vector3d()));
                                }
                                if (dynProp.PropertyName.ToUpper() == "DISTANCE_DRAWAREA_WIDTH")//stretch to be Width of selection + buffer %
                                {
                                    dynProp.Value = Convert.ToDouble(xDist * 1.1);
                                    //move to account for stretch change
                                }
                                //else if (dynProp.PropertyName.ToUpper() == "DISTANCE_NOTES")
                                //{ dynProp.Value = 1 * (.2075); }//need to calculate based on notes selected
                            }
                        }
                    }                    
                }
                trCurrent.Commit();
            }
        }
    }
}
