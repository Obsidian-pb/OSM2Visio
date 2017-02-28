using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace OSM2Visio.Code.DrawData
{

    class DrawINPPV_OSM
    {
        Microsoft.Office.Interop.Visio.Application VisioApp;
        System.Xml.XmlDocument Data;
        f_DrawProcess drawForm;

        //Переменные для работы
        System.Xml.XmlNodeList NodesList;
        //System.Xml.XmlNodeList NdList;
        System.Xml.XmlNodeList TdList;

        DrawTools.CoordRecatangle v_Box; // = new CoordRecatangle();

        Double x = 0;
        Double y = 0;
        Double XPos;
        Double YPos;

        Double InchInGradH;  //Количество дюймов в одном градусе долготы
        Double InchInGradV;  //Количество дюймов в одном градусе широты
        
        int i;

        public DrawINPPV_OSM(Microsoft.Office.Interop.Visio.Application _VisioApp, System.Xml.XmlDocument _Data,
            f_DrawProcess _drawForm, DrawTools.CoordRecatangle _v_Box)
        {
            VisioApp = _VisioApp;
            Data = _Data;
            drawForm = _drawForm;
            v_Box = _v_Box;

            //---Определяем количество дюймов в одном градусе долготы
            InchInGradH = DrawTools.GetInchesInGradH(v_Box);
            InchInGradV = DrawTools.GetInchesInGradV(v_Box);
        }



        /// <summary>
        /// Прока отрисовки ИНППВ
        /// </summary>
        public void DrawData()
        {
            //---Получаем список узлов с перечислением node
            NodesList = Data.SelectNodes("//node");
            //---Указываем максимальное значение процессбара
            drawForm.Text = "Расставляются водоисточники";
            drawForm.SetProgressbarMaximum(NodesList.Count);
            i = 0; // startValue;
            MessageBox.Show("Расставляются водоисточники");
            DrawTools.Coordinate pnt; pnt.x = 0; pnt.y = 0;
                    
            //---Перебираем все узлы node в списке NodeList
            foreach (System.Xml.XmlNode node in NodesList)
            {
                if (node.ChildNodes.Count > 0)
                {
                    TdList = node.SelectNodes("tag");  //список узлов с описанием
                    foreach (System.Xml.XmlNode Td in TdList)
                    {
                        if (Td.Attributes["k"].InnerText == "emergency")
                        {
                            //Получаем координаты точки где необходимо вставить ИНППВ
                            DrawTools.GetPosition(node.Attributes["id"].InnerText, ref Data, ref x, ref y);

                            //Получаем координату относительно края области (в дюймах - все в дюймах)
                            XPos = (x - v_Box.XY1.x);   YPos = (y - v_Box.XY1.y);
                            pnt.x = XPos * InchInGradH; pnt.y = YPos * InchInGradV;
                           
                            //Создаем новый ИНППВ, согласно указанным в node координатам
                            CreateEWS_OSM(ref VisioApp, TdList, Td.Attributes["v"].InnerText, pnt);
                        }
                    }

                }

                drawForm.SetProgressBarCurrentValue(i);
                i++;
                Application.DoEvents();
            }
        }


        //-------------------------------------------Вставка УГО ИНППВ--------------------------------------
        /// <summary>
        /// Прока вставляет значек ИНППВ
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateEWS_OSM(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            System.Xml.XmlNodeList TdList, String EWSType, DrawTools.Coordinate pnt)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            String tmpStr;

            switch (EWSType)
            {
                case "fire_hydrant":
                    //Вбрасываем новый ПГ
                    mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПГ"];
                    mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
                    shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

                    //Очищаем исходные данные
                    shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(" ");
                    shp.get_Cells("Prop.PGNumber").FormulaU = DrawTools.StringToFormulaForString(" ");

                    //Получаем данные о ПГ
                    foreach (System.Xml.XmlNode Td in TdList)
                    {
                        switch (Td.Attributes["k"].InnerText)
                        {
                            case "fire_hydrant:street":
                                shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText + " " + shp.get_Cells("Prop.PGAdress").FormulaU);
                                break;
                            case "fire_hydrant:housenumber":
                                shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(shp.get_Cells("Prop.PGAdress").FormulaU + " " + Td.Attributes["v"].InnerText);
                                break;
                            case "ref":
                                shp.get_Cells("Prop.PGNumber").FormulaU = DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                                break;
                            case "fire_hydrant:diameter":
                                tmpStr = Td.Attributes["v"].InnerText;
                                if (tmpStr.IndexOf("К") > 0)
                                    shp.get_Cells("Prop.PipeType").FormulaU = "INDEX(0, Prop.PipeType.Format)";
                                if (tmpStr.IndexOf("Т") > 0)
                                    shp.get_Cells("Prop.PipeType").FormulaU = "INDEX(1, Prop.PipeType.Format)";

                                //Убираем из строки ненужные символы
                                tmpStr = tmpStr.Replace("К", ""); tmpStr = tmpStr.Replace("Т", ""); tmpStr = tmpStr.Replace("-", "");

                                shp.get_Cells("Prop.PipeDiameter").FormulaU = DrawTools.StringToFormulaForString(tmpStr);
                                break;
                            case "fire_hydrant:pressure":
                                tmpStr = Td.Attributes["v"].InnerText;
                                switch (tmpStr)
                                {
                                    case "1":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("10");
                                        break;
                                    case "2":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("20");
                                        break;
                                    case "3":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("30");
                                        break;
                                    case "4":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("40");
                                        break;
                                    case "5":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("50");
                                        break;
                                    case "6":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("60");
                                        break;
                                    case "7":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("70");
                                        break;
                                    case "8":
                                        shp.get_Cells("Prop.Pressure").FormulaU = DrawTools.StringToFormulaForString("80");
                                        break;
                                    case "no":
                                        shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                                        shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                                        break;
                                }
                                break;
                            default:
                                break;
                        }
                        Application.DoEvents();
                    }
                    break;
                case "water_tank":
                    //Вбрасываем новый ПВ
                    mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПВ"];
                    mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
                    shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

                    //Получаем данные о ПВ
                    foreach (System.Xml.XmlNode Td in TdList)
                    {
                        switch (Td.Attributes["k"].InnerText)
                        {
                            case "water_tank:street":
                                shp.get_Cells("Prop.PWAdress").FormulaU = DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText + " " + shp.get_Cells("Prop.PWAdress").FormulaU);
                                break;
                            case "water_tank:housenumber":
                                shp.get_Cells("Prop.PWAdress").FormulaU = DrawTools.StringToFormulaForString(shp.get_Cells("Prop.PWAdress").FormulaU + " " + Td.Attributes["v"].InnerText);
                                break;
                            case "ref":
                                shp.get_Cells("Prop.PWNumber").FormulaU = DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                                break;
                            case "water_tank:volume":
                                tmpStr = Td.Attributes["v"].InnerText;
                                if (tmpStr.IndexOf("no") > 0)
                                {
                                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                                    tmpStr = "0";
                                }

                                shp.get_Cells("Prop.PWValue").FormulaU = DrawTools.StringToFormulaForString(tmpStr);
                                break;
                            default:
                                break;
                        }
                    }
                    break;
                default:
                    break;
            }
            return true;
        }


    }

}
