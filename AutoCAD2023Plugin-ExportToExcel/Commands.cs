// (C) Copyright 2022 by  
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Specialized;
using System.Collections;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;


[assembly: CommandClass(typeof(AutoCAD2023Plugin.Commands))]

namespace AutoCAD2023Plugin
{

    public class Commands
    {

        // Command exportexcel um Attribute eines Blocks im Dialog anzuzeigen und auf Knopfdruck exportieren zu können
        // NoUndoMarker : Keine Datenbank änderungen vorgenommen also keine Rücknehmbaren Änderungen in AutoCAD
        // Modal : Keine Benutzung während andere Commands laufen
        [CommandMethod("exportexcel", CommandFlags.NoUndoMarker | CommandFlags.Modal)]
        public void ExportToExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions options = new PromptEntityOptions("Bitte eine Blockreferenz auswählen");
            options.SetRejectMessage("\nAusgewähltes Objekt ist keine Blockreferenz");
            options.AddAllowedClass(typeof(BlockReference), true);

            PromptEntityResult result = ed.GetEntity(options);

            if (result.Status != PromptStatus.OK) return;

            using (Transaction transaction = db.TransactionManager.StartTransaction())
            {
                // Suchen nach den Attributen des Gewählten Objektes

                StringDictionary attrDict = new StringDictionary();
                BlockReference reference = (BlockReference)transaction.GetObject(result.ObjectId, OpenMode.ForRead);
                AttributeCollection attributeCollection = reference.AttributeCollection;

                foreach (ObjectId id in attributeCollection)
                {
                    // Einzelne Attribute herausnehmen und dem Dictionary hinzufügen

                    DBObject attrObj = transaction.GetObject(id, OpenMode.ForRead);
                    AttributeReference attrRef = attrObj as AttributeReference;

                    if (attrRef != null)
                    {
                        attrDict.Add(attrRef.Tag, attrRef.TextString);
                    }

                }

                // Keine Attribute im Block, Transaction abbrechen

                if(attrDict.Count == 0)
                {
                    ed.WriteMessage("\nBlock enthält keine Attribute");
                    transaction.Abort();
                    return;
                }

                // Control in der alle Attribute angezeigt werden 

                DataGridView gridView = new DataGridView
                {
                    ColumnCount = attrDict.Count + 1,
                    Name = reference.Name,
                    RowHeadersVisible = false,
                    Dock = DockStyle.Fill
                };
                gridView.Columns[0].Name = "Name";
                gridView.Rows[0].Cells[0].Value = reference.Name;
                int colNum = 1;

                // Bevölkerung der Tabelle/Control

                foreach (DictionaryEntry entry in attrDict)
                {
                    gridView.Columns[colNum].Name = (String)entry.Key;
                    gridView.Rows[0].Cells[colNum].Value = (String)entry.Value;
                    colNum++;
                }

                // Erstellen des Dialogs als extension von System.Windows.Forms

                ExportToExcel.ExportDialog form = new ExportToExcel.ExportDialog();

                form.AddControl(0, 0, 4, gridView);
                Application.ShowModalDialog(form);

                // Überprüfen auf DialogResult, wenn Ok wurde eine neue Datei erstellt

                if (form.DialogResult != DialogResult.OK)
                {
                    ed.WriteMessage("\nVorgang abgebrochen");
                    form.Dispose();
                    transaction.Abort();
                    return;
                }
                ed.WriteMessage("\nExcel-Datei im Ordner Dokumente gespeichert");
                form.Dispose();
                transaction.Commit();
            }
        }

    }

}





