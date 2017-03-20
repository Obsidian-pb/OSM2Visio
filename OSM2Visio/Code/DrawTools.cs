using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;
using System.Windows.Forms;

/*Класс для работы хранения служебных процедур отрисовки*/

namespace OSM2Visio
{
    static class DrawTools
    {
        public const double EARTH_RADIUS = 6371032;
        public const double PI = 3.141592654;
        public const double INCHINMETER = 39.3701;

        #region Структуры координат
        public struct Coordinate
        {
            public Double x, y;

            public void SetCrdnt(Double p1, Double p2)
            {
                x = p1;
                y = p2;
            }
        }
        public struct CoordRecatangle
        {
            public Coordinate XY1, XY2;
        }
        #endregion Структуры координат

        public enum INPPW_Types {PG, PW, MO, LO, NO, SO, Sk, Gr, Such, Ok, PK, PO, Bash, Pd, Pirs, nothing}

        /// <summary>
        /// Прока превращает входящую строку в строку формулы для вставки 
        /// в ячейки Visio. Заменяем все двойные кавычки(") парой двойных кавычек("")
        /// в начале и в конце строки.</summary>
        /// <param name="inputValue">Входящая строка</param>
        /// <returns>измененная строка, которая может быть программно
        /// назначена ячейке. Не может быть напрямую назначена ячейке, поскольку не имеет
        /// "=" в начале.</returns>
        static public string StringToFormulaForString(string inputValue)
        {
            string result = "";
            string quote = "\"";
            string quoteQuote = "\"\"";

            try
            {
                result = inputValue != null ? inputValue : String.Empty;

                // Заменяем (") на ("").
                result = result.Replace(quote, quoteQuote);

                // Добавляем ("") вокруг всей строки.
                result = quote + result + quote;
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }
            //MessageBox.Show(result);
            return result;
        }

        /// <summary>
        /// Функция корректно преобразует строку к типу Double
        /// </summary>
        /// <param name="a_Str"></param>
        /// <returns></returns>
        static public Double pf_StrToDbl(String a_Str)
        {
            try
            {
                return Convert.ToDouble(a_Str.Replace(".", ","));
            }
            catch (Exception err)
            {
                //MessageBox.Show(err.Message);
                return Convert.ToDouble(a_Str.Replace(",", "."));
                //throw;
            }

        }
        /// <summary>
        /// Функция возвращает True, если приложение имеет версию 16 или 15.
        /// </summary>
        /// <returns></returns>
        static public bool IsNewApp(String vers)
        {
            if (vers == "16,0" || vers == "16.0" || vers == "15,0" || vers == "15.0")
                return true;
            else
                return false;
        }


        #region Проки отрисовки OSM
        //static public void DrawRelation(System.Xml.XmlNode node, ref System.Xml.XmlDocument NodeDoc, DrawTools.CoordRecatangle v_Box,         
        //    double InchInGradH, double InchInGradV)
        //{
        //    foreach (System.Xml.XmlNode chldNode in node.ChildNodes)
        //    {
        //        if (chldNode.Name == "member")
        //        {
        //            if (chldNode.Attributes["type"].InnerText == "relation")
        //            {
        //                DrawRelation(chldNode, ref NodeDoc, v_Box, InchInGradH, InchInGradV);
        //            }
        //            if (chldNode.Attributes["type"].InnerText == "way")
        //            {
        //                DrawWay(chldNode, ref NodeDoc, v_Box, InchInGradH, InchInGradV);
        //            }
        //        }
        //    }
        //}

        //static public void DrawWay(System.Xml.XmlNode node, ref System.Xml.XmlDocument NodeDoc, DrawTools.CoordRecatangle v_Box,
        //    double InchInGradH, double InchInGradV)
        //{
        //    System.Xml.XmlNodeList NdList;

        //    Double x = 0;
        //    Double y = 0;
        //    Double XPos;
        //    Double YPos;

        //    int j = 0;

        //    NdList = node.SelectNodes("nd");  //список узлов с координатами точек
        //    //Массив для хранения точек для отрисовки зданий
        //    Array pnts = Array.CreateInstance(typeof(Double), NdList.Count * 2); ;  //-1

        //    //j = 0;
        //    //---Перебираем все узлы в списке NdList
        //    foreach (System.Xml.XmlNode Nd in NdList)
        //    {
        //        //PrB_DrawProcess.Value = i;

        //        GetPosition(Nd.Attributes["ref"].InnerText, ref NodeDoc, ref x, ref y);

        //        //Получаем координату относительно края области (в дюймах - все в дюймах)
        //        XPos = (x - v_Box.XY1.x);
        //        YPos = (y - v_Box.XY1.y);

        //        //Заполянем очередную точку в массиве
        //        pnts.SetValue(XPos * InchInGradH, j);
        //        pnts.SetValue(YPos * InchInGradV, j + 1);

        //        j = j + 2;
        //    }



        //}

        #endregion Проки отрисовки OSM



        #region Работа с координатами
        /// <summary>
        /// Прока передает в переменные данные об относительном положении точки на листе
        /// </summary>
        /// <param name="NodeID">ID узла</param>
        /// <param name="NodeDoc">Документ</param>
        /// <param name="x">координата X</param>
        /// <param name="y">координата Y</param>
        /// <returns></returns>
        static public Boolean GetPosition(String NodeID, ref System.Xml.XmlDocument NodeDoc,
            ref Double x, ref Double y)
        {
            //Переменные для работы
            System.Xml.XmlNodeList Nodes;

            //Получаем перечень узлов node
            Nodes = NodeDoc.SelectNodes("//node");

            //перебираем все элементы списка Nodes
            foreach (System.Xml.XmlNode node in Nodes)
            {

                if (node.Attributes["id"].InnerText == NodeID)
                {
                    x = pf_StrToDbl(node.Attributes["lon"].InnerText);
                    y = pf_StrToDbl(node.Attributes["lat"].InnerText);
                    return true;
                }
            }
            //Если ничего не найдено - делаем дополнительный запрос к сайту
            //GetPositionQuery NodeID, x, y;
            x = 0;
            y = 0;
            return false;
        }

        //internal static void GetPosition(string[] tempStrArr, ref double x, ref double y)
        //{
        //    throw new NotImplementedException();
        //}


        /// <summary>
        /// Функция возварщает количество дюймов в одном градусе долготы
        /// </summary>
        /// <param name="a_Box"></param>
        /// <returns></returns>
        static public Double GetInchesInGradH(CoordRecatangle a_Box)
        {
            double AngleDiff;
            double LengthDiff;

            //---Определяем разницу в углах X
            AngleDiff = a_Box.XY2.x - a_Box.XY1.x;
            //'---Определяем разницу в метрах -> (2 * PI * (Cos(Y1 * PI / 180) * EARTH_RADIUS) - Окружность земли на соответствующей широте
            LengthDiff = (AngleDiff / 360) * (2 * PI * (Math.Cos(a_Box.XY1.y * PI / 180) * EARTH_RADIUS));
            //'---Возвращаем дюймы в угле
            return (LengthDiff * INCHINMETER) / AngleDiff;
        }
        /// <summary>
        /// Функция возварщает количество дюймов в одном градусе широты
        /// </summary>
        /// <param name="a_Box"></param>
        /// <returns></returns>
        static public Double GetInchesInGradV(CoordRecatangle a_Box)
        {
            double AngleDiff;
            double LengthDiff;

            //---Определяем разницу в углах Y
            AngleDiff = a_Box.XY2.y - a_Box.XY1.y;
            //'---Определяем разницу в метрах
            LengthDiff = (AngleDiff / 360) * (2 * PI * EARTH_RADIUS);
            //'---Возвращаем дюймы в угле
            return (LengthDiff * INCHINMETER) / AngleDiff;
        }
        /// <summary>
        /// Прока возвраает линейные размеры прямоугольника
        /// </summary>
        /// <param name="a_Box">Прямоугольник</param>
        /// <param name="HorLen">Горизонатльная длина (по параллели)</param>
        /// <param name="VertLen">Вертикальная длина (по меридиану)</param>
        static public Boolean GetSizes(CoordRecatangle a_Box, ref Double HorLen, ref Double VertLen)
        {
            double AngleDiff;

            try
            {
                //---Определяем разницу в углах X
                AngleDiff = a_Box.XY2.x - a_Box.XY1.x;
                //'---Определяем разницу в метрах -> (2 * PI * (Cos(Y1 * PI / 180) * EARTH_RADIUS) - Окружность земли на соответствующей широте
                HorLen = (AngleDiff / 360) * (2 * PI * (Math.Cos(a_Box.XY1.y * PI / 180) * EARTH_RADIUS));

                //---Определяем разницу в углах Y
                AngleDiff = a_Box.XY2.y - a_Box.XY1.y;
                //'---Определяем разницу в метрах
                VertLen = (AngleDiff / 360) * (2 * PI * EARTH_RADIUS);
                return true;
            }
            catch (Exception err)
            {
                //MessageBox.Show(err.Message);
                return false;
                throw;
            }
        }
        /// <summary>
        /// Функция получает DOM документ с данными из OSM и рамку с координатами,
        /// в которую соохарняет данные о границе выборки из OSM. Возвращает True, если функция отработала и False если нет
        /// </summary>
        /// <param name="Data">Документ XML OSM</param>
        /// <param name="a_Box">ссылка на прямоугольник</param>
        /// <returns></returns>
        static public Boolean pb_GetBoundBox(System.Xml.XmlDocument Data, ref CoordRecatangle a_Box)
        {
            System.Xml.XmlNode vCR_BoundsNode;
            try
            {
                vCR_BoundsNode = Data.SelectNodes("//bounds")[0];
                a_Box.XY1.SetCrdnt(DrawTools.pf_StrToDbl(vCR_BoundsNode.Attributes["minlon"].InnerText),
                        DrawTools.pf_StrToDbl(vCR_BoundsNode.Attributes["minlat"].InnerText));
                a_Box.XY2.SetCrdnt(DrawTools.pf_StrToDbl(vCR_BoundsNode.Attributes["maxlon"].InnerText),
                        DrawTools.pf_StrToDbl(vCR_BoundsNode.Attributes["maxlat"].InnerText));
                return true;
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        }

        static public Boolean checkForBox(Coordinate _pnt, CoordRecatangle _box)
        {
            //MessageBox.Show(_pnt.x.ToString() + ", " + _pnt.y.ToString() + " --- " + _box.XY1.x.ToString() + ", " + _box.XY1.y.ToString() + ", " + _box.XY2.x.ToString() + ", " + _box.XY2.y.ToString());
            return (_pnt.x > _box.XY1.x && _pnt.x < _box.XY2.x && _pnt.y > _box.XY1.y && _pnt.y < _box.XY2.y);
        }
        #endregion Работа с координатами

        #region Работа с фигурами Visio
        /// <summary>
        /// Прока связывает ячейку двух фигур 
        /// </summary>
        /// <param name="a_shpFace"></param>
        /// <param name="a_shpBckgnd"></param>
        /// <param name="a_CellName"></param>
        static public void LinkShapes(ref Visio.Shape a_shpFace, ref Visio.Shape a_shpBckgnd, string a_CellName)
        {
            try
            {
                Visio.Cell ShpCell;

                ShpCell = a_shpBckgnd.get_Cells(a_CellName);
                ShpCell.FormulaForceU = "GUARD(" + a_shpFace.NameID + "!" + a_CellName + ")";
            }
            catch (Exception err)
            {
                //Ячейка ссылвется на несуществующую
                MessageBox.Show(err.Message);
                //throw;
            }

        }
        /// <summary>
        /// Копирует формулу из указанной ячейки фигуры 
        /// </summary>
        /// <param name="a_shpOrigin"></param>
        /// <param name="a_shpDescendent"></param>
        /// <param name="a_CellName"></param>
        static public void CopyCellFormula(ref Visio.Shape a_shpOrigin, ref Visio.Shape a_shpDescendent, string a_CellName)
        {
            try
            {
                    a_shpDescendent.get_Cells(a_CellName).FormulaU =
                        a_shpOrigin.get_Cells(a_CellName).FormulaU;                    
            }
            catch (Exception err)
            {
                //Ячейка ссылвется на несуществующую
                MessageBox.Show(err.Message);
                //throw;
            }
        }

        /// <summary>
        /// Прока обращает фигуру нарисованную при помощи метода PolylineTo
        /// в фигуру нарисованную при помощи метода LineTo
        /// </summary>
        /// <param name="VisioApp">Активное приложение Visio</param>
        /// <param name="shp">Фигура</param>
        static public void PolyLineToLine(ref Visio.Shape shp)
        {
            string PolyLineString;
            int i;
            short rowIndex = 2;

            try
            {
                //0 - Если фигура и так LineTo - выходим без изменений
                if (shp.get_RowType(10, 2) == 139)   //visSectionFirstComponent = 10, visTagPolylineTo = 193
                    return;

                //1 - сохраняем стартовые значения X.1 Y.1 для их использования в заключении
                string X1 = shp.get_Cells("Geometry1.X1").Formula;    //.get_ResultStr(0);
                string Y1 = shp.get_Cells("Geometry1.Y1").Formula;    //.get_ResultStr(0);

                //2 - получить строку с описанием линии
                PolyLineString = shp.get_Cells("Geometry1.A2").get_ResultStr(0) + ";";
                PolyLineString = PolyLineString.Replace("POLYLINE(", ""); PolyLineString = PolyLineString.Replace(")", "");

                //3 - Копируем значения вторйо строки в первую   
                //Отслеживаем циклическую ссылку
                if (shp.get_Cells("Geometry1.X2").Formula.Contains("Geometry1") || shp.get_Cells("Geometry1.Y2").Formula.Contains("Geometry1"))
                {
                    shp.get_Cells("Geometry1.X1").Formula = shp.get_Cells("Geometry1.X2").get_ResultStr(0);
                    shp.get_Cells("Geometry1.Y1").Formula = shp.get_Cells("Geometry1.Y2").get_ResultStr(0);
                }
                else
                {
                    shp.get_Cells("Geometry1.X1").Formula = shp.get_Cells("Geometry1.X2").Formula; //.get_ResultStr(0);
                    shp.get_Cells("Geometry1.Y1").Formula = shp.get_Cells("Geometry1.Y2").Formula; //.get_ResultStr(0);
                }


                //4 - Удаляем вторую строки
                shp.DeleteRow(10, 2);

                //5 - ОСНОВНАЯ  - создаем и заполняем строки по PolyLineString, в обратном порядке
                for (i = GetItemsCount(PolyLineString); i > 2; i -= 2)
                {
                    //5-1 Создаем новую строку
                    shp.AddRow(10, rowIndex, 0);
                    shp.set_RowType(10, rowIndex, 139);

                    //5-2 Заполняем для нее свойства
                    shp.get_CellsSRC(10, rowIndex, 0).Formula = "Width*" + GetStrByIndex(PolyLineString, i - 1).ToString();
                    shp.get_CellsSRC(10, rowIndex, 1).Formula = "Height*" + GetStrByIndex(PolyLineString, i).ToString();

                    rowIndex++;
                }

                //6 - создаем последнюю строку
                //6-1 Создаем новую строку
                shp.AddRow(10, rowIndex, 0);
                shp.set_RowType(10, rowIndex, 139);

                //6-2 Заполняем для нее свойства
                shp.get_CellsSRC(10, rowIndex, 0).Formula = X1;
                shp.get_CellsSRC(10, rowIndex, 1).Formula = Y1;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                throw;
            }
        }
        /// <summary>
        /// Прока обращает фигуру нарисованную при помощи метода PolylineTo
        /// в фигуру нарисованную при помощи метода LineTo
        /// ДЛЯ VISO 2016
        /// </summary>
        /// <param name="VisioApp">Активное приложение Visio</param>
        /// <param name="shp">Фигура</param>
        static public void PolyLineToLine16(ref Visio.Shape shp)
        {
            string PolyLineString;
            int i;
            short rowIndex = 2;
            double[] startShapeData = new double[4];
            string[] cellsFormulas = new string[2];

            try
            {
                //0 - Если фигура и так LineTo - выходим без изменений
                if (shp.get_RowType(10, 2) == 139 || shp.get_RowType(10, 2) == 239)
                    return;

                //0++ - сохраняем стартовое состояние фигуры
                startShapeData[0] = shp.get_Cells("Width").get_Result(Microsoft.Office.Interop.Visio.tagVisUnitCodes.visMeters);
                startShapeData[1] = shp.get_Cells("Height").get_Result(Microsoft.Office.Interop.Visio.tagVisUnitCodes.visMeters);
                startShapeData[2] = shp.get_Cells("PinX").get_Result(Microsoft.Office.Interop.Visio.tagVisUnitCodes.visMeters);
                startShapeData[3] = shp.get_Cells("PinY").get_Result(Microsoft.Office.Interop.Visio.tagVisUnitCodes.visMeters);
                cellsFormulas[0] = shp.get_Cells("Geometry1.X2").Formula;
                cellsFormulas[1] = shp.get_Cells("Geometry1.Y2").Formula;

                //1 - сохраняем стартовые значения X.1 Y.1 для их использования в заключении
                string X1 = shp.get_Cells("Geometry1.X1").Formula;    //.get_ResultStr(0);
                string Y1 = shp.get_Cells("Geometry1.Y1").Formula;    //.get_ResultStr(0);

                //2 - получить строку с описанием линии
                PolyLineString = shp.get_Cells("Geometry1.A2").get_ResultStr(0) + ";";
                PolyLineString = PolyLineString.Replace("POLYLINE(", ""); PolyLineString = PolyLineString.Replace(")", "");

                //3 - Копируем значения второй строки в первую   
                //Отслеживаем циклическую ссылку
                //Для 2016 - не требуется!

                //4 - Удаляем вторую строки
                shp.DeleteRow(10, 2);
                ////4++ - Восстанавливаем высоту фигуры
                shp.set_RowType((short)Microsoft.Office.Interop.Visio.tagVisSectionIndices.visSectionFirstComponent, (short)1, (short)138);
                shp.get_Cells("Geometry1.X1").Formula = cellsFormulas[0];
                shp.get_Cells("Geometry1.Y1").Formula = cellsFormulas[1];
                shp.get_Cells("Width").Formula = startShapeData[0].ToString() + " m";
                shp.get_Cells("Height").Formula = startShapeData[1].ToString() + " m";
                //MessageBox.Show("123");

                //5 - ОСНОВНАЯ  - создаем и заполняем строки по PolyLineString, в обратном порядке
                for (i = GetItemsCount(PolyLineString); i > 2; i -= 2)
                {
                    //5-1 Создаем новую строку
                    shp.AddRow(10, rowIndex, 0);
                    shp.set_RowType(10, rowIndex, 139);
                    shp.get_Cells("Width").Formula = startShapeData[0].ToString() + " m";
                    shp.get_Cells("Height").Formula = startShapeData[1].ToString() + " m";

                    //5-2 Заполняем для нее свойства
                    //MessageBox.Show("Width*" + GetStrByIndex(PolyLineString, i - 1).ToString() + " Height*" + GetStrByIndex(PolyLineString, i).ToString());
                    shp.get_CellsSRC(10, rowIndex, 0).Formula = "Width*" + GetStrByIndex(PolyLineString, i - 1).ToString();
                    shp.get_CellsSRC(10, rowIndex, 1).Formula = "Height*" + GetStrByIndex(PolyLineString, i).ToString();

                    rowIndex++;
                }

                //6 - создаем последнюю строку
                //6-1 Создаем новую строку
                shp.AddRow(10, rowIndex, 0);
                shp.set_RowType(10, rowIndex, 139);

                //6-2 Заполняем для нее свойства
                shp.get_CellsSRC(10, rowIndex, 0).Formula = X1;
                shp.get_CellsSRC(10, rowIndex, 1).Formula = Y1;

                //7 - Восстанавливаем исходные свойства фигуры
                //---Собственно восстановление
                //MessageBox.Show(cellsFormulas[0] + " : " + cellsFormulas[1]);
                shp.get_Cells("Geometry1.X1").Formula = cellsFormulas[0];
                shp.get_Cells("Geometry1.Y1").Formula = cellsFormulas[1];
                //---Повторно задаем значения для второй строки
                shp.get_Cells("Geometry1.X2").Formula = "Width*" + GetStrByIndex(PolyLineString, GetItemsCount(PolyLineString) - 1).ToString();
                shp.get_Cells("Geometry1.Y2").Formula = "Height*" + GetStrByIndex(PolyLineString, GetItemsCount(PolyLineString)).ToString();

                shp.get_Cells("Width").Formula = startShapeData[0].ToString() + " m";
                shp.get_Cells("Height").Formula = startShapeData[1].ToString() + " m";
                shp.get_Cells("PinX").Formula = startShapeData[2].ToString() + " m";
                shp.get_Cells("PinY").Formula = startShapeData[3].ToString() + " m";
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                throw;
            }
        }

        /// <summary>
        /// Прока возвращает Double значение строки с разделителем ";"
        /// ПЕРЕРАБОТАТЬ
        /// </summary>
        /// <param name="str"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        static public Double GetStrByIndex(string str, int index)
        {
            int i;
            int count = 0;
            int pos1 = 0;
            int pos2 = str.Length;
            string DblStr;


            if (GetItemsCount(str) == 0)
                return 0;
            try
            {
                for (i = 0; i < str.Length; i++)
                {
                    if (str.Substring(i, 1) == ";")
                    {
                        pos2 = i;
                        count++;
                        if (count == index)
                        {
                            DblStr = str.Substring(pos1, pos2 - pos1);
                            return DrawTools.pf_StrToDbl(DblStr);
                        }
                        pos1 = i + 1;
                    }
                }
                return 0;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return 0;
                //throw;
            }


        }

        static public int GetItemsCount(string str)
        {
            int i;
            int count = 0;

            for (i = 0; i < str.Length; i++)
            {
                if (str.Substring(i, 1) == ";")
                {
                    count++;
                }
            }
            return count;
        }

        #endregion Работа с фигурами Visio

        #region Работа со строками
        /// <summary>
        /// Функция возвращает подстроку с содержимым параметра
        /// </summary>
        /// <param name="description"></param>
        /// <param name="substr"></param>
        /// <returns></returns>
        static public string GetSubstringFromDescription(string description, string substr)
        {
            int pos1=0; int pos2=0;

            try
            {
                pos1 = description.IndexOf(substr) + substr.Length;
                for (int i = pos1; i < description.Length-5; i++)
			    {
			        if (description.Substring(i, 4)=="<br>")
                    {
                        pos2 = i;
                        break;
                    }
			    }
                return description.Substring(pos1, pos2 - pos1);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "";
                //throw;
            }
        }


        #endregion Работа со строками


    }
}
