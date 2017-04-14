using System;
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
        public byte pipeType;
        public int pipeDiameter;
        public int value;
        public byte PACount;
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

            int s = 0;

            try
            {
                //---Подключаемся к БД
                loadDb();
                if (checkCon())
	            {
                    connectionEWS.Open();
                    connectionCommand.CommandText = "SELECT EWS_GeoCoord_X, EWS_GeoCoord_Y, Streets.Street_Name, EWS_Building, EWS_Type, EWS_Number, EWS_PipeType, EWS_Diameter_COD, ";
                    connectionCommand.CommandText += "EWS_Value, EWS_PACount, EWS_Status,  EWS_PKDiameter ";
                    connectionCommand.CommandText += "FROM Streets LEFT JOIN EWSs ON Streets.Street_ID = EWSs.EWS_Street_COD ";
                    connectionCommand.CommandText += "WHERE (EWSs.EWS_GeoCoord_X<>0) AND (EWSs.EWS_GeoCoord_Y<>0);";
                    OleDbDataReader dbReader = connectionCommand.ExecuteReader();

                    //---Указываем максимальное значение процессбара
                    i = 0;
                    while (dbReader.Read()) i++; dbReader.Close();
                    drawForm.SetProgressbarMaximum(i);
                    dbReader = connectionCommand.ExecuteReader();

                    drawForm.Text = "Расставляются водоисточники";
                    i = 0; // startValue;
                    MessageBox.Show("Расставляются водоисточники");
                    DrawTools.Coordinate pnt; pnt.x = 0; pnt.y = 0;

                    i = 0;
                    //Читаем строки из запроса и на их основе формируем данные для отрисовки ИНППВ
                    while (dbReader.Read())
                    {
                        //Если форма закрыта - выходим
                        if (!drawForm.IsGoon) return;

                        //---Получаем данные из записи
                        //Получаем координаты точки где необходимо вставить ИНППВ
                        x = dbReader.GetDouble(0);  //"EWS_GeoCoord_X"
                        y = dbReader.GetDouble(1);  //"EWS_GeoCoord_Y"

                        DrawTools.Coordinate pnt2; pnt2.x = x; pnt2.y = y;

                        //Получаем координату относительно края области (в дюймах - все в дюймах)
                        XPos = (x - v_Box.XY1.x); YPos = (y - v_Box.XY1.y);
                        pnt.x = XPos * InchInGradH; pnt.y = YPos * InchInGradV;

                        
                        //Проверяем входит ли координата в прямоугольник карты
                        if (DrawTools.checkForBox(pnt2, v_Box))
                        {
                            ews_Data.address = GetStringData(dbReader, 2) + " " + GetStringData(dbReader, 3); s = 2; //---Адрес
                            ews_Data.INPPW_Type = GetInt32Data(dbReader, 4); s = 4; //---Тип ИНППВ
                            ews_Data.pipeType = GetByteData(dbReader, 6); s = 6; //---Водовод
                            ews_Data.pipeDiameter = GetInt16Data(dbReader, 7); s = 7;
                            ews_Data.number = GetStringData(dbReader, 5); s = 5; //---Номер
                            ews_Data.value = GetInt32Data(dbReader, 8); s = 8; //---Общий показатель (объем воды)
                            ews_Data.PACount = GetByteData(dbReader, 9); s = 9; //---Кол-во ПА
                            ews_Data.status = GetInt32Data(dbReader, 10); s = 10; //---Статус
                            ews_Data.PKDiameter = GetInt32Data(dbReader, 11); s = 11; //---Статус
                            //Создаем новый ИНППВ, согласно указанным в координатам, и передаем в него собранные данные
                            CreateEWS_EWS(ref VisioApp, pnt, ews_Data);
                        }
                        
                        i++;
                        drawForm.SetProgressBarCurrentValue(i);
                        Application.DoEvents();
                    }

                    //Закрываем соединение
                    connectionEWS.Close();
                    CloseDB();
	            }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + " s =" + s + " i =" + i);
                CloseDB();
                //throw;
            }
        }


        #region -------------------------------------------Работа с БД--------------------------------------------
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
        private string GetStringData(OleDbDataReader _Reader, int fieldIndex)
        {
            try {return _Reader.GetString(fieldIndex); }
            catch (Exception) { return ""; }
        }
        private int GetInt16Data(OleDbDataReader _Reader, int fieldIndex)
        {
            try { return _Reader.GetInt16(fieldIndex); }
            catch (Exception) { return 0; }
        }
        private int GetInt32Data(OleDbDataReader _Reader, int fieldIndex)
        {
            try { return _Reader.GetInt32(fieldIndex); }
            catch (Exception) { return 0; }
        }
        private byte GetByteData(OleDbDataReader _Reader, int fieldIndex)
        {
            try { return _Reader.GetByte(fieldIndex); }
            catch (Exception) { return 0; }
        }
        #endregion -------------------------------------------Работа с БД--------------------------------------------


        #region -------------------------------------------Вставка УГО ИНППВ--------------------------------------
        /// <summary>
        /// Прока вставляет значек ИНППВ
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateEWS_EWS(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;

            switch (ews_Data.INPPW_Type)
            {
                case 0:   //ПГ
                    //Вбрасываем новый ПГ
                    shp = DropNewPG(pnt, ews_Data);
                    break;
                case 1:   //ПВ
                    //Вбрасываем новый ПВ
                    shp = DropNewPW(pnt, ews_Data);
                    break;
                case 2:  //ПК
                    //Вбрасываем новый ПК
                    shp = DropNewPK(pnt, ews_Data);
                    break;
                case 3:  //Пирс
                    //Вбрасываем новый Пирс
                    shp = DropNewPirs(pnt, ews_Data);
                    break;
                case 4:  //Башня
                    //Вбрасываем новую Башню
                    shp = DropNewBash(pnt, ews_Data);
                    break;
                case 5:  //Колодец
                    //Вбрасываем новый Колодец
                    shp = DropNewKolodec(pnt, ews_Data);
                    break;
                default:
                    break;
            }

            return true;
        }
        #endregion -------------------------------------------Вставка УГО ИНППВ--------------------------------------


        #region Служебные функции

        private string GetTypePipe(int typePipe)
        {
            try
            {
                if (typePipe == 0)
                    return "Кольцевой";
                else
                    return "Тупиковый";
            }
            catch (Exception)
            {
                return "Кольцевой";
                //throw;
            }
        }
        private string GetDiameterPipe(int diameterCODPipe)
        {
            try
            {
                //Подключаемся к БД
                OleDbCommand connCmnd = new OleDbCommand();
                connCmnd.Connection = connectionEWS;
                //Формируем запрос о типах диаметров водоводов
                connCmnd.CommandText = "SELECT Diameters.* FROM Diameters;";
                //Читаем данные
                OleDbDataReader dbReader = connCmnd.ExecuteReader();

                while (dbReader.Read())
                {
                    if (dbReader.GetInt32(0) == diameterCODPipe) return dbReader.GetString(1);
                }

                dbReader = null;
                connCmnd = null;
                return "150";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "150";
                //throw;
            }
        }


        #endregion Служебные функции

        #region Проки вбрасывания новых фигур
        /// <summary>
        /// Функция всавки новой фигуры ПГ
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewPG(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПГ"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try            
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.PGNumber").FormulaU = DrawTools.StringToFormulaForString(ews_Data.number);
                shp.get_Cells("Prop.PGAdress").FormulaU = DrawTools.StringToFormulaForString(ews_Data.address);
                shp.get_Cells("Prop.PipeType").FormulaU = DrawTools.StringToFormulaForString(GetTypePipe(ews_Data.pipeType));
                shp.get_Cells("Prop.PipeDiameter").FormulaU = DrawTools.StringToFormulaForString(GetDiameterPipe(ews_Data.pipeDiameter));
                if (ews_Data.status==1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);                
                return shp;
            }
        }
        /// <summary>
        /// Функция всавки новой фигуры ПВ
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewPW(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПВ"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.PWNumber").FormulaU = DrawTools.StringToFormulaForString(ews_Data.number);
                shp.get_Cells("Prop.PWAdress").FormulaU = DrawTools.StringToFormulaForString(ews_Data.address);
                //shp.get_Cells("Prop.PipeType").FormulaU = DrawTools.StringToFormulaForString(GetTypePipe(ews_Data.pipeType));
                shp.get_Cells("Prop.PWValue").FormulaU = DrawTools.StringToFormulaForString(ews_Data.value.ToString());
                shp.get_Cells("Prop.SetsCount").FormulaU = DrawTools.StringToFormulaForString(ews_Data.PACount.ToString());
                if (ews_Data.status == 1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);                
                return shp;
            }
        }
        /// <summary>
        /// Функция всавки новой фигуры ПК
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewPK(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["ПК"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.PKNumber").FormulaU = DrawTools.StringToFormulaForString(ews_Data.number);
                shp.get_Cells("Prop.PKNumber").FormulaU = DrawTools.StringToFormulaForString(ews_Data.PKDiameter.ToString());
                if (ews_Data.status == 1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                return shp;
            }
        }
        /// <summary>
        /// Функция всавки новой фигуры Пирса
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewPirs(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Пирс"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.SetsCount").FormulaU = DrawTools.StringToFormulaForString(ews_Data.PACount.ToString());
                if (ews_Data.status == 1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                return shp;
            }
        }
        /// <summary>
        /// Функция всавки новой фигуры Башни
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewBash(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Башня"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.WTAdress").FormulaU = DrawTools.StringToFormulaForString(ews_Data.address);
                shp.get_Cells("Prop.WTValue").FormulaU = DrawTools.StringToFormulaForString(ews_Data.value.ToString());
                if (ews_Data.status == 1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                return shp;
            }
        }
        /// <summary>
        /// Функция всавки новой фигуры Колодца
        /// </summary>
        /// <param name="pnt">координаты ИНППВ</param>
        /// <param name="ews_Data">Объект данных ИНППВ</param>
        /// <returns>shp - фигура ПГ</returns>
        private Visio.Shape DropNewKolodec(DrawTools.Coordinate pnt, EWS_Data ews_Data)
        {
            Visio.Shape shp;
            Visio.Master mstr;

            mstr = VisioApp.Documents["Водоснабжение.vss"].Masters["Колодец"];
            mstr.Shapes[1].get_Cells("EventDrop").FormulaU = "";  //Отключаем событие вброса для данной фигуры
            shp = VisioApp.ActivePage.Drop(mstr.Shapes[1], pnt.x, pnt.y);

            try
            {
                //Уазываем данные ПГ
                shp.get_Cells("Prop.WTAdress").FormulaU = DrawTools.StringToFormulaForString(ews_Data.address);
                if (ews_Data.status == 1)
                {
                    shp.get_Cells("LineColor").FormulaU = DrawTools.StringToFormulaForString("2");
                    shp.get_Cells("Char.Color").FormulaU = DrawTools.StringToFormulaForString("2");
                }
                return shp;
            }
            catch (Exception)
            {
                return shp;
            }
        }


        #endregion Проки вбрасывания новых фигур

    }
}
