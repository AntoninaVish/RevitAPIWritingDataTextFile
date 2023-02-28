using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIWritingDataTextFile
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //создаем переменную, которая будет собирать данные по помещению
            string roomInfo = string.Empty;//переменная в которую эти данные будут записываться

            var rooms = new FilteredElementCollector(doc) //собираем сами помещения с помощью FilteredElementCollector
            .OfCategory(BuiltInCategory.OST_Rooms) //собираем по категории
            .Cast<Room>() //элементы, которые получаем преобразовываем в помещения
            .ToList();//преобразовываем в список

            //заполняем данный параметр roominfo, проходимся в цикле по каждому помещению
            foreach (Room room in rooms)
            {
                //извлекаем имя помещения, для этого создаем отдельную переменную, с помощью get_parameter заходим
                //в строенный параметри находим встроенный параметр ROOM_NAME, AsString - имя помещения выводится в строку
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                
                //заполняем переменную roomInfo новыми данными, \t - это знак табуляции, чтобы выводилось в разном формате txt, xl
                roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}";//имя, номер, площадь и все выводиться
                                                                                           //с новой строки
            }

            //сохраняем файл, напр., на рабочем столе
            //получаем путь к папке рабочего стола: desktopPath - путь к рабочему столу; GetFolderPath - получить путь к рабочей папке;
            //Desktop- рабочий стол
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //указываем полный путь сохранения имено текстового файла, для этого соеденяем desktopPath и название файла "roomImfo.csv"
            string csvPath = Path.Combine(desktopPath, "roomInfo.csv");

            //запись всего текста в файл
            File.WriteAllText(csvPath, roomInfo);


            return Result.Succeeded;
        }
    }
}
