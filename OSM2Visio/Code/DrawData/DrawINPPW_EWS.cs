﻿using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections.Generic;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace OSM2Visio
{
    public struct EWS_Data
    {
        public int INPPW_Type;
        public string address;
        public int pipeType;
        public int pipeDiameter;
        public int value;
        public int PACount;
        public int status;
        public string number;
        public int PKDiameter;
    }

    class DrawINPPW_EWS
    {
        Microsoft.Office.Interop.Visio.Application VisioApp;
        f_DrawProcess drawForm;

        //Переменные для работы
        public string dataBasePath;
        //База данных
        OleDbConnection connectionEWS = new OleDbConnection();
        OleDbCommand connectionCommand = new OleDbCommand();

        DrawTools.CoordRecatangle v_Box; 

        Double x = 0;
        Double y = 0;
        Double XPos;
        Double YPos;

        Double InchInGradH;  //Количество дюймов в одном градусе долготы
        Double InchInGradV;  //Количество дюймов в одном градусе широты

        public DrawINPPW_EWS(Microsoft.Office.Interop.Visio.Application _VisioApp, string _dataBasePath,
            f_DrawProcess _drawForm, DrawTools.CoordRecatangle _v_Box)
        {
            VisioApp = _VisioApp;
            dataBasePath = _dataBasePath;
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
            EWS_Data ews_Data;

            try
            {
                //---Подключаемся к БД
                loadDb();
                if (checkCon())
	            {
                    connectionEWS.Open();
                    connectionCommand.CommandText = "SELECT EWS_GeoCoord_X, EWS_GeoCoord_Y, Streets.Street_Name, EWS_Building, EWS_Type, EWS_Number, EWS_PipeType, EWS_Diameter_COD ";
                    connectionCommand.CommandText += "EWS_Value, EWS_PACount, EWS_Status,  EWS_PKDiameter ";
                    connectionCommand.CommandText += "FROM Streets LEFT JOIN EWSs ON Streets.Street_ID = EWSs.EWS_Street_COD ";
                    connectionCommand.CommandText += "WHERE (EWSs.EWS_GeoCoord_X<>0) AND (EWSs.EWS_GeoCoord_Y<>0);";
                    MessageBox.Show(connectionCommand.CommandText);
                    OleDbDataReader dbReader = connectionCommand.ExecuteReader();

                        //---Указываем максимальное значение процессбара
                        //foreach (System.Xml.XmlNode node in dbReader.)
                        //{
                        //    if (node.Name == "Placemark") i++;
                        //}

                    drawForm.Text = "Расставляются водоисточники";
                    i = 0; // startValue;
                    MessageBox.Show("Расставляются водоисточники");
                    DrawTools.Coordinate pnt; pnt.x = 0; pnt.y = 0;


                    //Читаем строки из запроса и на их основе формируем данные для отрисовки ИНППВ
                    while (dbReader.Read())
                    {
                        //MessageBox.Show(Convert.ToString(dbReader["EWS_GeoCoord_X"]));

                        //---Получаем данные из записи
                        //Получаем координаты точки где необходимо вставить ИНППВ
                        //x = DrawTools.pf_StrToDbl(dbReader["EWS_GeoCoord_X"].ToString());
                        //y = DrawTools.pf_StrToDbl(dbReader["EWS_GeoCoord_Н"].ToString());
                        x = dbReader.GetDouble(0);  //"EWS_GeoCoord_X"
                        y = dbReader.GetDouble(1);  //"EWS_GeoCoord_Y"

                        DrawTools.Coordinate pnt2; pnt2.x = x; pnt2.y = y;

                        //Получаем координату относительно края области (в дюймах - все в дюймах)
                        XPos = (x - v_Box.XY1.x); YPos = (y - v_Box.XY1.y);
                        pnt.x = XPos * InchInGradH; pnt.y = YPos * InchInGradV;

                        //Проверяем входит ли координата в прямоугольник карты
                        if (DrawTools.checkForBox(pnt2, v_Box))
                        {
                            ews_Data.address = dbReader.GetString(2) + " " + dbReader.GetString(3); //---Адрес
                            ews_Data.INPPW_Type = dbReader.GetInt32(4);  //---Тип ИНППВ
                            ews_Data.pipeType = dbReader.GetInt32(6); //---Водовод
                            ews_Data.pipeDiameter = dbReader.GetInt32(7);
                            ews_Data.number = dbReader.GetString(5); //---Номер
                            ews_Data.value = dbReader.GetInt32(8); //---Общий показатель (объем воды)
                            ews_Data.PACount = dbReader.GetInt32(9); //---Кол-во ПА
                            ews_Data.status = dbReader.GetInt32(10); //---Статус
                            ews_Data.PKDiameter = dbReader.GetInt32(11); //---Статус

                            //Создаем новый ИНППВ, согласно указанным в координатам, и передаем в него собранные данные
                            //CreateEWS_ESU(ref VisioApp, pnt, INPPW_Type, description, condition, caption);
                        }

                        //            drawForm.SetProgressBarCurrentValue(i);
                        //            i++;
                        //        }







                    }

                    //Закрываем соединение
                    connectionEWS.Close();
                    CloseDB();
	            }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                CloseDB();
                //throw;
            }

            







            //try
            //{
            //    string[] ewsData = new String[10];
            //    //---Получаем список узлов с перечислением node
            //    System.Xml.XmlNode DocNode = Data.ChildNodes.Item(1).ChildNodes.Item(0);
            //    //NodesList = DocNode.SelectNodes("/Placemark");
            //    //---Указываем максимальное значение процессбара
            //    foreach (System.Xml.XmlNode node in DocNode.ChildNodes)
            //    {
            //        if (node.Name == "Placemark") i++;
            //    }
            //    drawForm.SetProgressbarMaximum(i); i = 0;

            
        }


        //-------------------------------------------Работа с БД--------------------------------------------
        public void loadDb()
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + this.dataBasePath + ";Persist Security Info=False;";
            connectionEWS.ConnectionString = connectionString;
            connectionCommand.Connection = connectionEWS;
        }
        public void CloseDB()
        {
            connectionEWS.Dispose();
        }

        public bool checkCon()
        {
            try
            {
                connectionEWS.Open();
                connectionEWS.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }



        //-------------------------------------------Вставка УГО ИНППВ--------------------------------------
        /// <summary>
        /// Прока вставляет значек ИНППВ
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateEWS_EWS(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            DrawTools.Coordinate pnt, DrawTools.INPPW_Types INPPW_Type, string description,
            string condition, string caption)
        {
            Visio.Shape shp;
            string numberINPPW;
            string address;
            bool state;

            address = DrawTools.GetSubstringFromDescription(description, "Улица (наименование объекта): ");
            numberINPPW = GetNumberINPPW(caption);
            state = GetStateINPPW(description);

            switch (INPPW_Type)
            {
                case DrawTools.INPPW_Types.PG:
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.PW:
                    //Вбрасываем новый ПВ
                    shp = DropNewPW(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.MO:
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.LO:
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.NO:
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.SO:
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Sk:
                    //Вбрасываем новый колодец
                    shp = DropNewSK(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Gr:
                    //Вбрасываем новый водоем
                    shp = DropNewPW(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Such:
                    //Вбрасываем новый ПК - так как специального символа для Сухотруба нет
                    shp = DropNewPK(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Ok:
                    break;
                case DrawTools.INPPW_Types.PK:
                    //Вбрасываем новый ПК
                    shp = DropNewPK(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.PO:
                    break;
                case DrawTools.INPPW_Types.Bash:
                    //Вбрасываем новую башню
                    shp = DropNewBash(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Pd:
                    //Вбрасываем новый пирс
                    shp = DropNewPirs(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                case DrawTools.INPPW_Types.Pirs:
                    //Вбрасываем новый пирс
                    shp = DropNewPirs(pnt, caption, address, numberINPPW, state);
                    AddCommonData(shp, description);
                    break;
                default:
                    break;
            }

            return true;
        }

        #region Работа с фигурами
        /// <summary>
        /// Метод копирует дополнительные сведения о ИППВ
        /// </summary>
        /// <param name="shp"></param>
        /// <param name="_description"></param>
        private void AddCommonData(Visio.Shape shp, string _description)
        {
            string commonData = _description.Replace("<br>", "\n");
            shp.get_Cells("Prop.Common").FormulaU = DrawTools.StringToFormulaForString(commonData);
        }
        #endregion Работа с фигурами

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
        private string GetDIameterPKFromCaption(string caption)
        {
            try
            {
                int pos1=50;
                if(caption.IndexOf("ПК-")>=0)
                    pos1 = caption.IndexOf("ПК-")+3;
                if(caption.IndexOf("Сух-")>=0)
                    pos1 = caption.IndexOf("Сух-")+4;
                return caption.Substring(pos1);
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
                return "50";
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
        /// <summary>
        /// Прока получает из названия ПВ его объем
        /// </summary>
        /// <param name="caption">заголовок ИНППВ</param>
        /// <returns>объем водоема</returns>
        private string GetValuePW(string caption)
        {
            try
            {
                int pos1 = 0;
                if (caption.IndexOf("ПВ-") >= 0)
                    pos1 = caption.IndexOf("ПВ-") + 3;
                if (caption.IndexOf("Гр-") >= 0)
                    pos1 = caption.IndexOf("Гр-") + 3;
                double value = Double.Parse(caption.Substring(pos1))/1000;
                return value.ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "5";
                //throw;
            }
        }
        private string GetValueWT(string caption)
        {
            try
            {
                int pos1  = caption.IndexOf("Баш-") + 4;
                double value = Double.Parse(caption.Substring(pos1));
                return value.ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "50";
                //throw;
            }
        }

        private string GetValuePA(string caption)
        {
            try
            {
                int pos1 = 0;
                if (caption.IndexOf("Пд-") >= 0)
                    pos1 = caption.IndexOf("Пд-") + 3;
                if (caption.IndexOf("Пирс-") >= 0)
                    pos1 = caption.IndexOf("Пирс-") + 5;
                return caption.Substring(pos1);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "2";
                //throw;
            }
        }

        private DrawTools.INPPW_Types GetTypeINPPW(string caption)
        {
            if (caption.IndexOf("ПГ")>=0) {return DrawTools.INPPW_Types.PG;}
            if (caption.IndexOf("ПВ")>=0) {return DrawTools.INPPW_Types.PW;}
            if (caption.IndexOf("МО")>=0) {return DrawTools.INPPW_Types.MO;}
            if (caption.IndexOf("ЛО") >= 0) { return DrawTools.INPPW_Types.LO; }
            if (caption.IndexOf("НО") >= 0) { return DrawTools.INPPW_Types.NO; }
            if (caption.IndexOf("СО") >= 0) { return DrawTools.INPPW_Types.SO; }
            if (caption.IndexOf("Ск") >= 0) { return DrawTools.INPPW_Types.Sk; }
            if (caption.IndexOf("Гр") >= 0) { return DrawTools.INPPW_Types.Gr; }
            if (caption.IndexOf("Сух") >= 0) { return DrawTools.INPPW_Types.Such; }
            if (caption.IndexOf("Ок") >= 0) { return DrawTools.INPPW_Types.Ok; }
            if (caption.IndexOf("ПК") >= 0) { return DrawTools.INPPW_Types.PK; }
            if (caption.IndexOf("ПО") >= 0) { return DrawTools.INPPW_Types.PO; }
            if (caption.IndexOf("Баш") >= 0) { return DrawTools.INPPW_Types.Bash; }
            if (caption.IndexOf("Пд") >= 0) { return DrawTools.INPPW_Types.Pd; }
            if (caption.IndexOf("Пирс") >= 0) { return DrawTools.INPPW_Types.Pirs; }
            return DrawTools.INPPW_Types.nothing;
        }


        #endregion Служебные функции

        #region Проки вбрасывания новых фигур
        /// <summary>
        /// Функция всавки новой фигуры ПГ
        /// </summary>
        /// <param name="pnt"></param>
        /// <param name="caption"></param>
        /// <param name="address"></param>
        /// <param name="numberINPPW"></param>
        /// <param name="state"></param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewPG(DrawTools.Coordinate pnt, string caption, 
             string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string typePG;
            string diameter;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПГ"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try            
            {
                //Дополнительные сведения о ПГ
                typePG = GetTypePG(caption);
                diameter = GetDIameterFromCaption(caption);
                //Уазываем данные ПГ
                shp.get_Cells("Prop.PGNumber").FormulaU = DrawTools.StringToFormulaForString(numberINPPW);
                shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(address);
                shp.get_Cells("Prop.PipeType").FormulaU = DrawTools.StringToFormulaForString(typePG);
                shp.get_Cells("Prop.PipeDiameter").FormulaU = DrawTools.StringToFormulaForString(diameter);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);                
                return shp;
                //throw;
            }
        }



        private Visio.Shape DropNewSK(DrawTools.Coordinate pnt, string caption,
             string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Колодец"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные колодца
                shp.get_Cells("Prop.Adress").FormulaU = DrawTools.StringToFormulaForString(address);
                shp.get_Cells("Prop.About").FormulaU = DrawTools.StringToFormulaForString(caption);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return shp;
                //throw;
            }
        }
        /// <summary>
        /// Функция вбрасывает фигуру пожарного крана и присваивает ей значения
        /// </summary>
        /// <param name="pnt"></param>
        /// <param name="caption"></param>
        /// <param name="address"></param>
        /// <param name="numberINPPW"></param>
        /// <param name="state"></param>
        /// <returns></returns>
        private Visio.Shape DropNewPK(DrawTools.Coordinate pnt, string caption,
             string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string diameter;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПК"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Дополнительные сведения о ПК
                diameter = GetDIameterPKFromCaption(caption);
                //Уазываем данные ПК
                shp.get_Cells("Prop.PKNumber").FormulaU = DrawTools.StringToFormulaForString(numberINPPW);
                shp.get_Cells("Prop.PKDiameter").FormulaU = DrawTools.StringToFormulaForString(diameter);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return shp;
                //throw;
            }
        }

        private Visio.Shape DropNewPW(DrawTools.Coordinate pnt, string caption,
             string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string valuePW;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПВ"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Дополнительные сведения о ПВ
                valuePW = GetValuePW(caption);
                //Уазываем данные ПГ
                shp.get_Cells("Prop.PWNumber").FormulaU = DrawTools.StringToFormulaForString(numberINPPW);
                shp.get_Cells("Prop.PWAdress").FormulaU = DrawTools.StringToFormulaForString(address);
                shp.get_Cells("Prop.PWValue").FormulaU = DrawTools.StringToFormulaForString(valuePW);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return shp;
                //throw;
            }
        }

        private Visio.Shape DropNewBash(DrawTools.Coordinate pnt, string caption,
             string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string valueWT;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Башня"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Дополнительные сведения о ПВ
                valueWT = GetValueWT(caption);
                //Уазываем данные ПГ
                shp.get_Cells("Prop.WTAdress").FormulaU = DrawTools.StringToFormulaForString(address);
                shp.get_Cells("Prop.WTValue").FormulaU = DrawTools.StringToFormulaForString(valueWT);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return shp;
                //throw;
            }
        }

        private Visio.Shape DropNewPirs(DrawTools.Coordinate pnt, string caption,
                     string address, string numberINPPW, bool state)
        {
            Visio.Shape shp;
            Visio.Master mstr;
            string valuePA;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Пирс"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Дополнительные сведения о Пирсе
                valuePA = GetValuePA(caption);
                //Уазываем данные ПГ
                shp.get_Cells("Prop.SetCount").FormulaU = DrawTools.StringToFormulaForString(valuePA);
                if (!state)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return shp;
                //throw;
            }
        }  

        #endregion Проки вбрасывания новых фигур

    }
}