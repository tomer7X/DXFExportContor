using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace DXFExportContor
{
    public class Class1
    {
        [CommandMethod("HelloWorld")]
        public void HelloWorld()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage("\nHello World from DXFExportContor!");
        }

        [CommandMethod("ExplodeArrays")]
        public void ExplodeArrays()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt the user to select objects
            PromptSelectionResult selResult = ed.GetSelection();
            if (selResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo objects selected.");
                return;
            }

            SelectionSet selSet = selResult.Value;
            int explodedCount = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace =
                    (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent is not BlockReference blkRef) continue;

                    // Check if this block reference is an associative array
                    if (!IsAssociativeArray(blkRef, tr)) continue;

                    blkRef.UpgradeOpen();

                    // Recursively explode until we reach primitive geometry
                    List<Entity> primitives = [];
                    if (!DeepExplode(blkRef, primitives))
                    {
                        ed.WriteMessage("\nFailed to explode an array object.");
                        continue;
                    }

                    // Add all primitive entities to the current space
                    foreach (Entity primitive in primitives)
                    {
                        currentSpace.AppendEntity(primitive);
                        tr.AddNewlyCreatedDBObject(primitive, true);
                    }

                    // Remove the original array block reference
                    blkRef.Erase();
                    explodedCount++;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nExploded {explodedCount} array(s).");
        }

        private static bool DeepExplode(Entity entity, List<Entity> results)
        {
            DBObjectCollection explodedObjects = new DBObjectCollection();
            try
            {
                entity.Explode(explodedObjects);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                return false;
            }

            foreach (DBObject obj in explodedObjects)
            {
                if (obj is BlockReference nestedBlkRef)
                {
                    // Keep exploding nested block references
                    if (!DeepExplode(nestedBlkRef, results))
                    {
                        // If it can't be exploded further, keep it as-is
                        results.Add(nestedBlkRef);
                    }
                }
                else if (obj is Entity ent)
                {
                    results.Add(ent);
                }
            }

            return true;
        }


        private static bool IsAssociativeArray(BlockReference blkRef, Transaction tr)
        {
            // Method 1: Check if the block reference itself has an extension
            // dictionary with ACAD_ASSOCNETWORK
            if (!blkRef.ExtensionDictionary.IsNull)
            {
                DBDictionary extDict =
                    (DBDictionary)tr.GetObject(blkRef.ExtensionDictionary, OpenMode.ForRead);

                if (extDict.Contains("ACAD_ASSOCNETWORK"))
                    return true;
            }

            // Method 2: Check the block definition (BlockTableRecord) for
            // an extension dictionary with ACAD_ASSOCNETWORK.
            // AutoCAD stores the associative network on the BTR, not the reference.
            BlockTableRecord btr =
                (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);

            if (!btr.ExtensionDictionary.IsNull)
            {
                DBDictionary btrExtDict =
                    (DBDictionary)tr.GetObject(btr.ExtensionDictionary, OpenMode.ForRead);

                if (btrExtDict.Contains("ACAD_ASSOCNETWORK"))
                    return true;
            }

            // Method 3: Check persistent reactors on the block reference.
            // Associative arrays attach AssocDependency objects as reactors.
            ObjectIdCollection reactorIds = blkRef.GetPersistentReactorIds();
            if (reactorIds != null)
            {
                foreach (ObjectId reactorId in reactorIds)
                {
                    if (!reactorId.IsValid || reactorId.IsErased) continue;

                    DBObject reactor = tr.GetObject(reactorId, OpenMode.ForRead);
                    string rxClassName = reactor.GetRXClass().Name;

                    if (rxClassName.Contains("AssocDependency") ||
                        rxClassName.Contains("AssocArray"))
                        return true;
                }
            }

            // Method 4: Check if the block name follows the anonymous array
            // naming pattern (e.g., "*A1", "*A2", etc.)
            if (btr.Name.StartsWith("*A", System.StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        [CommandMethod("Tomer")]
        public void Tomer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt user to select the grid block
            PromptEntityOptions peo = new("\nSelect the grid block:");
            peo.SetRejectMessage("\nMust be a block reference.");
            peo.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo block selected.");
                return;
            }

            const double cellWidth = 6000.0;
            const double cellHeight = 3000.0;
            const int cols = 10;

            using Transaction tr = db.TransactionManager.StartTransaction();

            // Get the selected block's bounding box to determine the grid origin
            BlockReference gridBlock =
                (BlockReference)tr.GetObject(per.ObjectId, OpenMode.ForRead);

            Extents3d blockExtents;
            try
            {
                blockExtents = gridBlock.GeometricExtents;
            }
            catch
            {
                ed.WriteMessage("\nCould not get block extents.");
                return;
            }

            // Grid origin is the top-left corner of the block
            double originX = blockExtents.MinPoint.X;
            double originY = blockExtents.MaxPoint.Y;

            // Collect all text entities on layer "P_Names" in model space
            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

            List<(Point3d Position, string Text)> pNamesTexts = [];
            foreach (ObjectId entId in modelSpace)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!ent.Layer.Equals("P_Names", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (ent is DBText dbText)
                    pNamesTexts.Add((dbText.Position, dbText.TextString));
                else if (ent is MText mText)
                    pNamesTexts.Add((mText.Location, mText.Contents));
            }

            // Iterate: start at top-left, go left to right (10 cols),
            // then move down one row. Stop when a cell has no P_Names text.
            int cellCount = 0;
            for (int row = 0; ; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    double cellMinX = originX + col * cellWidth;
                    double cellMaxY = originY - row * cellHeight;
                    double cellMaxX = cellMinX + cellWidth;
                    double cellMinY = cellMaxY - cellHeight;

                    // Find text whose position falls inside this cell
                    string foundText = null;
                    foreach (var (pos, text) in pNamesTexts)
                    {
                        if (pos.X >= cellMinX && pos.X <= cellMaxX &&
                            pos.Y >= cellMinY && pos.Y <= cellMaxY)
                        {
                            foundText = text;
                            break;
                        }
                    }

                    if (foundText == null)
                    {
                        ed.WriteMessage(
                            $"\nCell [row {row + 1}, col {col + 1}]: No text on layer P_Names. Stopping.");
                        ed.WriteMessage($"\nProcessed {cellCount} cell(s) total.");
                        tr.Commit();
                        return;
                    }

                    cellCount++;
                    ed.WriteMessage($"\nCell [row {row + 1}, col {col + 1}]: {foundText}");
                }
            }
        }
    }
}