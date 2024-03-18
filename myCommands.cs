using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.Geometry;
using System.Diagnostics;

// Данный класс содержит методы для использования в AutoCAD через командную строку
public class TextCoordinatesExporter
{
    // Команда для AutoCAD, которую можно вызвать через консоль команд
    // Например, введя команду EXPORTTEXTCOORDS
    [CommandMethod("EXPORTTEXTCOORDS")]
    public void ExportTextCoordinatesToFile()
    {
        // Получаем текущий документ и его редактор
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;
        Editor ed = acDoc.Editor;

        try
        {
            // Выбираем объекты в документе
            PromptSelectionResult selection = ed.GetSelection();
            if (selection.Status != PromptStatus.OK) return;
            SelectionSet set = selection.Value;

            // Создаем файл для записи
            string fileName = "TextCoordinates.csv";
            using (StreamWriter sw = new StreamWriter(fileName, false, Encoding.UTF8))
            {
                // Записываем заголовки столбцов
                sw.WriteLine("Text String, X, Y, Z");

                // Запускаем транзакцию
                using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Перебираем все выбранные объекты
                    foreach (SelectedObject obj in set)
                    {
                        if (obj != null)
                        {
                            // Открываем объект для чтения
                            Entity ent = trans.GetObject(obj.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent is DBText || ent is MText)
                            {
                                Point3d position;
                                string text;

                                if (ent is DBText dbText)
                                {
                                    position = dbText.Position;
                                    text = dbText.TextString;
                                }
                                else
                                {
                                    MText mText = ent as MText;
                                    position = mText.Location;
                                    text = mText.Text;
                                }

                                // Записываем данные в файл
                                sw.WriteLine($"\"{text}\", {position.X}, {position.Y}, {position.Z}");
                            }
                        }
                    }
                    // Завершаем транзакцию
                    trans.Commit();
                }
            }

            // Сообщаем пользователю об успешной записи файла
            ed.WriteMessage($"\nЭкспорт координат текста выполнен успешно. Файл сохранен: {fileName}");

            // Получаем полный путь к файлу
            string fullPath = Path.GetFullPath(fileName);

            // Открываем папку с файлом и выделяем файл
            Process.Start("explorer.exe", $"/select, \"{fullPath}\"");
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
            // В случае ошибки выводим сообщение
            ed.WriteMessage($"\nОшибка: {ex.Message}");
        }
    }
}