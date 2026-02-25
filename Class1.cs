using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace DXFExportContor
{
    public class ContourExporter
    {
        private const double CellWidth = 10000.0;
        private const double CellHeight = 5000.0;
        private const int Cols = 10;

        [CommandMethod("NHdxfCreator")]
        public void SeperatePanelsToDXF()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectId gridBlockId = PromptForGridBlock(ed);
            if (gridBlockId.IsNull) return;

            ObjectId contourObjectId = PromptForContourObject(ed);
            if (contourObjectId.IsNull) return;

            string dwgFolder = System.IO.Path.GetDirectoryName(doc.Name);
            string dwgName = System.IO.Path.GetFileNameWithoutExtension(doc.Name);
            string outputFolder = System.IO.Path.Combine(dwgFolder, dwgName + "_DXF");

            using Transaction tr = db.TransactionManager.StartTransaction();

            string contourLayer = ((Entity)tr.GetObject(contourObjectId, OpenMode.ForRead)).Layer;
            ed.WriteMessage($"\nContour layer: {contourLayer}");

            Extents3d blockExtents = GetBlockExtents(tr, gridBlockId);
            if (blockExtents.MinPoint == blockExtents.MaxPoint)
            {
                ed.WriteMessage("\nCould not get block extents.");
                return;
            }

            double originX = blockExtents.MinPoint.X;
            double originY = blockExtents.MaxPoint.Y;

            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

            int explodedCount = ExplodeArraysInArea(modelSpace, tr,
                blockExtents.MinPoint.X, blockExtents.MinPoint.Y,
                blockExtents.MaxPoint.X, blockExtents.MaxPoint.Y);

            if (explodedCount > 0)
                ed.WriteMessage($"\nExploded {explodedCount} array(s) in the block area.");

            var pNamesTexts = CollectTextOnLayer(modelSpace, tr, "P_Names");
            var contourEntities = CollectEntitiesOnLayer(modelSpace, tr, contourLayer);

            if (!System.IO.Directory.Exists(outputFolder))
                System.IO.Directory.CreateDirectory(outputFolder);

            int cellCount = 0;
            for (int row = 0; ; row++)
            {
                for (int col = 0; col < Cols; col++)
                {
                    double cellMinX = originX + col * CellWidth;
                    double cellMaxY = originY - row * CellHeight;
                    double cellMaxX = cellMinX + CellWidth;
                    double cellMinY = cellMaxY - CellHeight;

                    string name = FindTextInCell(pNamesTexts, cellMinX, cellMinY, cellMaxX, cellMaxY);
                    if (name == null)
                    {
                        ed.WriteMessage(
                            $"\nCell [row {row + 1}, col {col + 1}]: No text on layer P_Names. Stopping.");
                        ed.WriteMessage($"\nExported {cellCount} DXF file(s) to: {outputFolder}");
                        tr.Commit();
                        return;
                    }

                    List<ObjectId> cellEntityIds = FindEntitiesInCell(
                        contourEntities, cellMinX, cellMinY, cellMaxX, cellMaxY);

                    if (cellEntityIds.Count == 0)
                    {
                        ed.WriteMessage(
                            $"\nCell [row {row + 1}, col {col + 1}]: {name} - No contour entities, skipping.");
                        cellCount++;
                        continue;
                    }

                    string dxfPath = System.IO.Path.Combine(outputFolder, SanitizeFileName(name) + ".dxf");
                    ExportEntitiesToDxf(db, cellEntityIds, dxfPath);

                    cellCount++;
                    ed.WriteMessage(
                        $"\nCell [row {row + 1}, col {col + 1}]: {name} - Exported {cellEntityIds.Count} entities.");
                }
            }
        }

        #region User Prompts

        private static ObjectId PromptForGridBlock(Editor ed)
        {
            PromptEntityOptions peo = new("\nSelect the grid block:");
            peo.SetRejectMessage("\nMust be a block reference.");
            peo.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo block selected.");
                return ObjectId.Null;
            }
            return per.ObjectId;
        }

        private static ObjectId PromptForContourObject(Editor ed)
        {
            PromptEntityOptions peo = new("\nSelect an object on the contour layer:");
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo object selected.");
                return ObjectId.Null;
            }
            return per.ObjectId;
        }

        #endregion

        #region Data Collection

        private static Extents3d GetBlockExtents(Transaction tr, ObjectId blockId)
        {
            BlockReference gridBlock = (BlockReference)tr.GetObject(blockId, OpenMode.ForRead);
            try { return gridBlock.GeometricExtents; }
            catch { return new Extents3d(); }
        }

        private static List<(Point3d Position, string Text)> CollectTextOnLayer(
            BlockTableRecord modelSpace, Transaction tr, string layerName)
        {
            List<(Point3d, string)> results = [];
            foreach (ObjectId entId in modelSpace)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (ent is DBText dbText)
                    results.Add((dbText.Position, dbText.TextString));
                else if (ent is MText mText)
                    results.Add((mText.Location, mText.Contents));
            }
            return results;
        }

        private static List<(ObjectId Id, Point3d Center)> CollectEntitiesOnLayer(
            BlockTableRecord modelSpace, Transaction tr, string layerName)
        {
            List<(ObjectId, Point3d)> results = [];
            foreach (ObjectId entId in modelSpace)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                if (!ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                    continue;

                Extents3d ext;
                try { ext = ent.GeometricExtents; }
                catch { continue; }

                Point3d center = new(
                    (ext.MinPoint.X + ext.MaxPoint.X) / 2.0,
                    (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0,
                    0);

                results.Add((entId, center));
            }
            return results;
        }

        #endregion

        #region Cell Lookup

        private static string FindTextInCell(
            List<(Point3d Position, string Text)> texts,
            double minX, double minY, double maxX, double maxY)
        {
            foreach (var (pos, text) in texts)
            {
                if (pos.X >= minX && pos.X <= maxX &&
                    pos.Y >= minY && pos.Y <= maxY)
                    return text;
            }
            return null;
        }

        private static List<ObjectId> FindEntitiesInCell(
            List<(ObjectId Id, Point3d Center)> entities,
            double minX, double minY, double maxX, double maxY)
        {
            List<ObjectId> result = [];
            foreach (var (id, center) in entities)
            {
                if (center.X >= minX && center.X <= maxX &&
                    center.Y >= minY && center.Y <= maxY)
                    result.Add(id);
            }
            return result;
        }

        #endregion

        #region DXF Export

        private static void ExportEntitiesToDxf(
            Database sourceDb, List<ObjectId> entityIds, string dxfPath)
        {
            using Database destDb = new(true, true);
            destDb.CloseInput(true);

            ObjectIdCollection ids = new([.. entityIds]);
            IdMapping idMap = new();

            using (Transaction destTr = destDb.TransactionManager.StartTransaction())
            {
                BlockTableRecord destModelSpace = (BlockTableRecord)destTr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(destDb), OpenMode.ForWrite);

                sourceDb.WblockCloneObjects(
                    ids, destModelSpace.ObjectId, idMap, DuplicateRecordCloning.Replace, false);

                ZoomExtents(destDb, destTr, destModelSpace);
                destTr.Commit();
            }

            destDb.DxfOut(dxfPath, 16, DwgVersion.AC1015);
        }

        private static void ZoomExtents(
            Database db, Transaction tr, BlockTableRecord modelSpace)
        {
            Extents3d totalExtents = new();
            bool hasExtents = false;

            foreach (ObjectId entId in modelSpace)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                try
                {
                    if (!hasExtents)
                    {
                        totalExtents = ent.GeometricExtents;
                        hasExtents = true;
                    }
                    else
                    {
                        totalExtents.AddExtents(ent.GeometricExtents);
                    }
                }
                catch { }
            }

            if (!hasExtents) return;

            ViewportTableRecord vp = (ViewportTableRecord)tr.GetObject(
                db.CurrentViewportTableRecordId, OpenMode.ForWrite);

            double centerX = (totalExtents.MinPoint.X + totalExtents.MaxPoint.X) / 2.0;
            double centerY = (totalExtents.MinPoint.Y + totalExtents.MaxPoint.Y) / 2.0;
            double width = (totalExtents.MaxPoint.X - totalExtents.MinPoint.X) * 1.1;
            double height = (totalExtents.MaxPoint.Y - totalExtents.MinPoint.Y) * 1.1;

            vp.CenterPoint = new Point2d(centerX, centerY);
            vp.Height = height;
            vp.Width = width;
        }

        #endregion

        #region Array Explode

        private static int ExplodeArraysInArea(
            BlockTableRecord space, Transaction tr,
            double minX, double minY, double maxX, double maxY)
        {
            List<BlockReference> arraysToExplode = [];
            foreach (ObjectId entId in space)
            {
                Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent is not BlockReference blkRef) continue;
                if (!IsAssociativeArray(blkRef, tr)) continue;

                Extents3d ext;
                try { ext = blkRef.GeometricExtents; }
                catch { continue; }

                if (ext.MaxPoint.X < minX || ext.MinPoint.X > maxX ||
                    ext.MaxPoint.Y < minY || ext.MinPoint.Y > maxY)
                    continue;

                arraysToExplode.Add(blkRef);
            }

            int count = 0;
            foreach (BlockReference blkRef in arraysToExplode)
            {
                blkRef.UpgradeOpen();

                List<Entity> primitives = [];
                if (!DeepExplode(blkRef, primitives))
                    continue;

                foreach (Entity primitive in primitives)
                {
                    space.AppendEntity(primitive);
                    tr.AddNewlyCreatedDBObject(primitive, true);
                }

                blkRef.Erase();
                count++;
            }
            return count;
        }

        private static bool DeepExplode(Entity entity, List<Entity> results)
        {
            DBObjectCollection exploded = new();
            try { entity.Explode(exploded); }
            catch (Autodesk.AutoCAD.Runtime.Exception) { return false; }

            foreach (DBObject obj in exploded)
            {
                if (obj is BlockReference nested)
                {
                    if (!DeepExplode(nested, results))
                        results.Add(nested);
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
            if (!blkRef.ExtensionDictionary.IsNull)
            {
                DBDictionary extDict =
                    (DBDictionary)tr.GetObject(blkRef.ExtensionDictionary, OpenMode.ForRead);
                if (extDict.Contains("ACAD_ASSOCNETWORK"))
                    return true;
            }

            BlockTableRecord btr =
                (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);

            if (!btr.ExtensionDictionary.IsNull)
            {
                DBDictionary btrExtDict =
                    (DBDictionary)tr.GetObject(btr.ExtensionDictionary, OpenMode.ForRead);
                if (btrExtDict.Contains("ACAD_ASSOCNETWORK"))
                    return true;
            }

            ObjectIdCollection reactorIds = blkRef.GetPersistentReactorIds();
            if (reactorIds != null)
            {
                foreach (ObjectId reactorId in reactorIds)
                {
                    if (!reactorId.IsValid || reactorId.IsErased) continue;
                    string rxName = tr.GetObject(reactorId, OpenMode.ForRead).GetRXClass().Name;
                    if (rxName.Contains("AssocDependency") || rxName.Contains("AssocArray"))
                        return true;
                }
            }

            if (btr.Name.StartsWith("*A", System.StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        #endregion

        #region Utilities

        private static string SanitizeFileName(string name)
        {
            char[] invalid = System.IO.Path.GetInvalidFileNameChars();
            foreach (char c in invalid)
                name = name.Replace(c, '_');
            return name.Trim();
        }

        #endregion
    }
}