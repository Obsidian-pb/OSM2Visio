using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace OSM2Visio.Code.DrawData
{
    class DrawINPPW_ESU
    {
        Microsoft.Office.Interop.Visio.Application VisioApp;
        System.Xml.XmlDocument Data;    //Здесь - документ KML - извлеченный из KMZ
        Object kmzFile;                 //Документ KMZ
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

        public DrawINPPW_ESU(Microsoft.Office.Interop.Visio.Application _VisioApp, System.Xml.XmlDocument _Data,
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
            int i=0;
            DrawTools.INPPW_Types INPPW_Type;
            string[] tempStrArr;
            string tempStr="";
            string description;
            string condition;
            string caption;

            try
            {
                string[] ewsData = new String[10];
                //---Получаем список узлов с перечислением node
                System.Xml.XmlNode DocNode = Data.ChildNodes.Item(1).ChildNodes.Item(0);
                //NodesList = DocNode.SelectNodes("/Placemark");
                //---Указываем максимальное значение процессбара
                foreach (System.Xml.XmlNode node in DocNode.ChildNodes)
                {
                    if (node.Name == "Placemark") i++;
                }
                drawForm.SetProgressbarMaximum(i); i = 0;
                drawForm.Text = "Расставляются водоисточники";
                //drawForm.SetProgressbarMaximum(NodesList.Count);
                i = 0; // startValue;
                MessageBox.Show("Расставляются водоисточники");
                DrawTools.Coordinate pnt; pnt.x = 0; pnt.y = 0;

                //---Перебираем все узлы node в списке NodeList
                //foreach (System.Xml.XmlNode node in NodesList)
                foreach (System.Xml.XmlNode node in DocNode.ChildNodes)
                {
                    if (node.Name == "Placemark")
                    {
                        //---Получаем данные из узла
                        //---Координаты точки
                        tempStrArr = node.ChildNodes.Item(3).FirstChild.InnerText.Split(',');
                        //Получаем координаты точки где необходимо вставить ИНППВ
                        x = DrawTools.pf_StrToDbl(tempStrArr[0]);
                        y = DrawTools.pf_StrToDbl(tempStrArr[1]);

//Проверяем входит ли координата в прямоугольник карты

                        //Получаем координату относительно края области (в дюймах - все в дюймах)
                        XPos = (x - v_Box.XY1.x); YPos = (y - v_Box.XY1.y);
                        pnt.x = XPos * InchInGradH; pnt.y = YPos * InchInGradV;

                        //---Описание
                        caption = node.ChildNodes.Item(0).InnerText;

                        //---Тип ИНППВ
                        //tempStr = node.ChildNodes.Item(0).InnerText;
                        tempStr = caption.Substring(2, 2);
                        #region Определяем тип ИНППВ
                        switch (tempStr)
                        {
                            case "ПГ":
                                INPPW_Type = DrawTools.INPPW_Types.PG;
                                break;
                            case "ПВ":
                                INPPW_Type = DrawTools.INPPW_Types.PW;
                                break;
                            case "МО":
                                INPPW_Type = DrawTools.INPPW_Types.MO;
                                break;
                            case "ЛО":
                                INPPW_Type = DrawTools.INPPW_Types.LO;
                                break;
                            case "НО":
                                INPPW_Type = DrawTools.INPPW_Types.NO;
                                break;
                            case "СО":
                                INPPW_Type = DrawTools.INPPW_Types.SO;
                                break;
                            case "Ск":
                                INPPW_Type = DrawTools.INPPW_Types.Sk;
                                break;
                            case "Гр":
                                INPPW_Type = DrawTools.INPPW_Types.Gr;
                                break;
                            case "Су":
                                INPPW_Type = DrawTools.INPPW_Types.Such;
                                break;
                            case "Ок":
                                INPPW_Type = DrawTools.INPPW_Types.Ok;
                                break;
                            case "ПК":
                                INPPW_Type = DrawTools.INPPW_Types.PK;
                                break;
                            case "ПО":
                                INPPW_Type = DrawTools.INPPW_Types.PO;
                                break;
                            case "Ба":
                                INPPW_Type = DrawTools.INPPW_Types.Bash;
                                break;
                            case "Пд":
                                INPPW_Type = DrawTools.INPPW_Types.Pd;
                                break;
                            case "Пирс":
                                INPPW_Type = DrawTools.INPPW_Types.Pirs;
                                break;
                            default:
                                INPPW_Type = DrawTools.INPPW_Types.nothing;
                                break;
                        }
                        #endregion Определяем тип ИНППВ

                        //---Описание
                        description = node.ChildNodes.Item(2).InnerText;
                        description = description.Substring(2, description.Length - 4);

                        //---Состояние
                        condition = node.ChildNodes.Item(1).InnerText;

                        //Создаем новый ИНППВ, согласно указанным в node координатам
                        CreateEWS_ESU(ref VisioApp, pnt, INPPW_Type, description, condition, caption);

                        drawForm.SetProgressBarCurrentValue(i);
                        i++;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                //throw;
            }
            
        }


        //-------------------------------------------Вставка УГО ИНППВ--------------------------------------
        /// <summary>
        /// Прока вставляет значек ИНППВ
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateEWS_ESU(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            DrawTools.Coordinate pnt, DrawTools.INPPW_Types INPPW_Type, string description,
            string condition, string caption)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string numberINPPW;
            string typePG;
            string diameter;
            string address;
            bool state;
            //string 

            address = DrawTools.GetSubstringFromDescription(description, "Улица (наименование объекта): ");
            numberINPPW = GetNumberINPPW(caption);
            state = GetStateINPPW(description);


            switch (INPPW_Type)
            {
                case DrawTools.INPPW_Types.PG:
                    //Вбрасываем новый ПГ
                    mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПГ"];
                    mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
                    shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);
                    //Дополнительные сведения о ПГ
                    typePG = GetTypePG(caption);                
                    diameter = GetDIameterFromCaption(caption);
                    //Уазываем данные ПГ
                    //MessageBox.Show(address + " ; " + numberINPPW + " ; " + state + " ; " + typePG + " ; " + diameter);
                    shp.get_Cells("Prop.PGNumber").FormulaU = DrawTools.StringToFormulaForString(numberINPPW);
                    shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(address);
                    shp.get_Cells("Prop.PipeType").FormulaU = DrawTools.StringToFormulaForString(typePG);
                    shp.get_Cells("Prop.PipeDiameter").FormulaU = DrawTools.StringToFormulaForString(diameter);
                    if (!state)
                    {
                        shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                        shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                    }
                    break;
                case DrawTools.INPPW_Types.PW:
                    break;
                case DrawTools.INPPW_Types.MO:
                    break;
                case DrawTools.INPPW_Types.LO:
                    break;
                case DrawTools.INPPW_Types.NO:
                    break;
                case DrawTools.INPPW_Types.SO:
                    break;
                case DrawTools.INPPW_Types.Sk:
                    break;
                case DrawTools.INPPW_Types.Gr:
                    break;
                case DrawTools.INPPW_Types.Such:
                    break;
                case DrawTools.INPPW_Types.Ok:
                    break;
                case DrawTools.INPPW_Types.PK:
                    break;
                case DrawTools.INPPW_Types.PO:
                    break;
                case DrawTools.INPPW_Types.Bash:
                    break;
                case DrawTools.INPPW_Types.Pd:
                    break;
                case DrawTools.INPPW_Types.Pirs:
                    break;
                default:
                    break;
            }

            //добавляем описание ИНППВ в ячейку фигуры и делаем команду меню видимой



            return true;
        }

        #region Служебные функции
        /// <summary>
        /// Функция возвращает диаметр водовода из строки заголовка ИНППВ в ЭСУ ППВ
        /// </summary>
        /// <param name="caption"></param>
        /// <returns></returns>
        private string GetDIameterFromCaption(string caption)
        {
            try
            {
                int pos1 = caption.IndexOf('-');
                return caption.Substring(pos1 + 1);
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
                return "150";
                //throw;
            }
        }
        private string GetNumberINPPW(string caption)
        {
            try
            {
                int pos1 = caption.IndexOf(' ');
                return caption.Substring(0, pos1);
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
                return " ";
                //throw;
            }
        }
        private bool GetStateINPPW(string description)
        {
            try
            {
                string state = DrawTools.GetSubstringFromDescription(description, "Техническое состояние: ");
                if (state == "Неисправен")
                    return false;
                else
                    return true;
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
                return true;
                //throw;
            }
            //Техническое состояние: 
        }
        private string GetTypePG(string caption)
        {
            try
            {
                if (caption.IndexOf("К-") > 0)
                    return "Кольцевой";
                else
                    return "Тупиковый";
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
                return "Кольцевой";
                //throw;
            }
        }

        #endregion Служебные функции



    }
}
