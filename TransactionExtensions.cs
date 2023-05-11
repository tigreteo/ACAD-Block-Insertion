using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace BlockInsertionV3
{
    public static class TransactionExtensions
    {
        // A simple extension method that aggregates the extents of any entities
        // from Through the interface author Kean Walmsley
        public static Extents3d GetExtents(this Transaction tr, SelectionSet ids)
        {
            var ext = new Extents3d();
            foreach (var id in ids.GetObjectIds())
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                { ext.AddExtents(ent.GeometricExtents); }
            }
            return ext;
        }
    }
}
