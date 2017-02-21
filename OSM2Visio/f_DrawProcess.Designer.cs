using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
using System.Windows.Forms;
//using System.Xml.Linq;
//using System.IO.Compression;
using System.IO;
using Ionic.Zip;
//using System.Array;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using OSM2Visio.Code.DrawData;
using OSM2Visio.Code;


namespace OSM2Visio
{
    partial class f_DrawProcess
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(f_DrawProcess));
            this.PrB_DrawProcess = new System.Windows.Forms.ProgressBar();
            this.B_OK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // PrB_DrawProcess
            // 
            this.PrB_DrawProcess.Location = new System.Drawing.Point(12, 12);
            this.PrB_DrawProcess.Name = "PrB_DrawProcess";
            this.PrB_DrawProcess.Size = new System.Drawing.Size(447, 36);
            this.PrB_DrawProcess.TabIndex = 0;
            // 
            // B_OK
            // 
            this.B_OK.Enabled = false;
            this.B_OK.Location = new System.Drawing.Point(367, 54);
            this.B_OK.Name = "B_OK";
            this.B_OK.Size = new System.Drawing.Size(92, 26);
            this.B_OK.TabIndex = 8;
            this.B_OK.Text = "Готово";
            this.B_OK.UseVisualStyleBackColor = true;
            this.B_OK.Click += new System.EventHandler(this.B_OK_Click);
            // 
            // f_DrawProcess
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 90);
            this.Controls.Add(this.B_OK);
            this.Controls.Add(this.PrB_DrawProcess);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "f_DrawProcess";
            this.Text = "Отрисовка данных";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar PrB_DrawProcess;
        private System.Windows.Forms.Button B_OK;

        //------------Проки отрисовки
        //---------------------------Внешние проки формы
        //Основная прока отрисовки зданий из OSM, Получает XMLDocument документ с данными из файла
        public void Pv_Draw(Microsoft.Office.Interop.Visio.Application VisioApp,
            System.Xml.XmlDocument Data , int INPPVSourceIndex, string EWSFilePath)
        {
            //Переменные для работы
            System.Xml.XmlNodeList NodesList;
            System.Xml.XmlNodeList NdList;
            System.Xml.XmlNodeList TdList;

            DrawTools.CoordRecatangle v_Box = new DrawTools.CoordRecatangle();

            //Dim Crdnt As Coordinate  '- доработать
            Double x = 0;
            Double y = 0;
            Double XPos;
            Double YPos;

            Double InchInGradH;  //Количество дюймов в одном градусе долготы
            Double InchInGradV;  //Количество дюймов в одном градусе широты

            Visio.Shape shp;

            Boolean NeedDelete; //Флаг необходимости удаления исходной фигуры

            int i;
            int j;
            short k;

            //---Собственно прока
            //---Показываем форму
            this.Show();

            //---Создаем слой "Здания"
            //VisioApp.ActiveWindow.Page.Layers.Add("Здания");

            //---Получаем рамку ограждающую выборку из OSM
            v_Box.XY1.SetCrdnt(0, 0);
            v_Box.XY2.SetCrdnt(0, 0);

            DrawTools.pb_GetBoundBox(Data, ref v_Box);

            //---Определяем количество дюймов в одном градусе долготы
            InchInGradH = DrawTools.GetInchesInGradH(v_Box);
            InchInGradV = DrawTools.GetInchesInGradV(v_Box);
            //MessageBox.Show("Размеры вычислены: Гор -> " + InchInGradH.ToString("R"));

            //---Определяем линейные размеры прямоугольника и приравниваем к нему рабочий лист
            SetSizeScale(ref VisioApp, v_Box);
            //---Увеличиваем картинку листа по размеру окна  Application.ActiveWindow.ViewFit = visFitPage
            VisioApp.ActiveWindow.ViewFit = (int)Visio.VisWindowFit.visFitPage;
            this.Focus();
            this.Top = 200; this.Left = 400;

            //---Получаем узел с перечислением Way
            NodesList = Data.SelectNodes("//way");
            //---Указываем максимальное значение процессбара
            this.Text = "Импортируются фигуры";
            this.PrB_DrawProcess.Maximum = NodesList.Count;
            i = 0;

            //---Перебираем все узлы way в списке NodeList
            foreach (System.Xml.XmlNode node in NodesList)
            {
                NdList = node.SelectNodes("nd");  //список узлов с координатами точек
                //Массив для хранения точек для отрисовки зданий
                Array pnts = Array.CreateInstance(typeof(Double), NdList.Count * 2); ;  //-1
                
                j = 0;
                //---Перебираем все узлы в списке NdList
                foreach (System.Xml.XmlNode Nd in NdList)
                {
                    this.PrB_DrawProcess.Value = i;

                    DrawTools.GetPosition(Nd.Attributes["ref"].InnerText, ref Data, ref x, ref y);
                    
                    //Получаем координату относительно края области (в дюймах - все в дюймах)
                    XPos = (x - v_Box.XY1.x);
                    YPos = (y - v_Box.XY1.y);
                    
                    //Заполянем очередную точку в массиве
                    pnts.SetValue(XPos * InchInGradH, j);
                    pnts.SetValue(YPos * InchInGradV, j + 1);
                    
                    j = j + 2;
                }
                //Рисуем фигуру по полученному массиву точек
                shp = VisioApp.ActivePage.DrawPolyline(ref pnts, 0);

                //Дописываем совйства фигуры
                TdList = node.SelectNodes("tag");
                k = 0;
                
                //Перебираем тэги "tag" и устанавлваем все свойства
                NeedDelete = false;
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    //В зависимости от того, что за объект - определяем его свойства 
                    //и необходимость удаления исходной геометрии
                    if (Td.Attributes["k"].InnerText == "building")
                    {
                        CreateCorrectBuilding(ref VisioApp, ref shp, TdList);
                        NeedDelete = true;
                    }
                    if (Td.Attributes["k"].InnerText == "highway")
                    {
                        CreateCorrectRoad(ref VisioApp, ref shp, TdList);
                        NeedDelete = true;
                    }
                    if (Td.Attributes["k"].InnerText == "landuse")
                    {
                        CreateCorrectLandUse(ref VisioApp, ref shp, TdList);
                        NeedDelete = false;
                    }
                    if (Td.Attributes["k"].InnerText == "leisure")
                    {
                        CreateCorrectLeisure(ref VisioApp, ref shp, TdList);
                        NeedDelete = false;
                    }
                    if (Td.Attributes["k"].InnerText == "barrier")
                    {
                        DrawTools.PolyLineToLine(ref shp);
                        CreateCorrectBorder(ref VisioApp, ref shp, TdList);
                        NeedDelete = false;
                    }
                    Application.DoEvents();

                    k++;
                }
                if (NeedDelete)
                    shp.Delete();
                i++;
            }

            //В зависимости от типа выбранного источника данных по ИНППВ выполняем импорт
            switch (INPPVSourceIndex)
            {
                case 0:   //Файл данных OSM
                    DrawINPPV_OSM INPPv_OSM = new DrawINPPV_OSM(VisioApp, Data, this, v_Box);
                    INPPv_OSM.DrawData();
                    INPPv_OSM = null;
                    break;
                case 1:  //Файл БД EWS
                    DrawINPPW_EWS INPPW_EWS = new DrawINPPW_EWS(VisioApp, EWSFilePath, this, v_Box);
                    INPPW_EWS.DrawData();
                    INPPW_EWS = null;
                    break;
                case 2:  //Файл строки подключения к БД

                    break;
                case 3:  //Файл данных ЭСУ ППВ
                    //Получаем документ XML со сведенями о ИНППВ
                    System.Xml.XmlDocument INPPW_Data = GetKML(EWSFilePath);
                    DrawINPPW_ESU INPPW_ESU = new DrawINPPW_ESU(VisioApp, INPPW_Data, this, v_Box);
                    INPPW_ESU.DrawData();
                    INPPW_ESU = null;
                    break;
                default:

                    break;
            }




            //Распределяем слои - здания вперед, территории назад
            LayersFix(ref VisioApp);

            //Отчет о завршении
            MessageBox.Show("Отрисовано " + i.ToString() + " объектов");
            this.B_OK.Enabled = true;

        }



        //------------------------Проки и функции отрисовки
        #region Проки и функции отрисовки
        /// <summary>
        /// Функция заменяет нарисованную из OSM фигуру на две фигуры дороги
        /// - лицевую часть и задний фон. Так же лицевой фигуре присваиваются свойства
        /// из узла "tag", описывающие доргоу
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="shp">Фигура</param>
        /// <param name="TdList">Перечень узлов со свойствами</param>
        /// <returns></returns>
        private Boolean CreateCorrectRoad(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            ref Visio.Shape shp, System.Xml.XmlNodeList TdList) 
        {
            try
            {
                Visio.Master BcgndMstr;
                Visio.Shape BcgndShp;
                Visio.Cell ShpCell;
                short i;
                
                //создаем покрытие (лицевую часть)
                BcgndMstr = VisioApp.Documents["План на местности.vss"].Masters["Дорога 2"];
                BcgndShp = VisioApp.ActivePage.Drop(BcgndMstr.Shapes[1], 0, 0);
                i = BcgndShp.get_RowCount(243);  //Определяем количество строк в секции visCustomProps

                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Width");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Height");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Angle");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "PinX");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "PinY");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "LocPinX");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "LocPinY");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.X1");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.Y1");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.X2");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.Y2");
                DrawTools.CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.A2");

                //Получаем все свойства для данного объекта
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    switch (Td.Attributes["k"].InnerText)
                    {
                        case "highway":
                            //определяем что за дорога и в соответствии с этим указываем ее ширину
                            switch (Td.Attributes["v"].InnerText)
                            {
                                case "trunk":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(8,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(0,Prop.RoadType.Format)";                                    
                                    break;
                                case "primary":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(6,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(1,Prop.RoadType.Format)";
                                    break;
                                case "secondary":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(4,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(1,Prop.RoadType.Format)";
                                    break;
                                case "tertiary":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(4,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(1,Prop.RoadType.Format)";
                                    break;
                                case "living_street":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(4,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(1,Prop.RoadType.Format)";
                                    break;
                                case "service":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(4,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(2,Prop.RoadType.Format)";
                                    break;  
                                case "residential":
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(4,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(2,Prop.RoadType.Format)";
                                    break;
                                default:
                                    BcgndShp.get_Cells("Prop.RoadWidth").FormulaU = "INDEX(2,Prop.RoadWidth.Format)";
                                    BcgndShp.get_Cells("Prop.RoadType").FormulaU = "INDEX(3,Prop.RoadType.Format)";
                                    break;
                            }
                            break;
                        case "name":
                            BcgndShp.get_Cells("Prop.Street").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                            break;
                        case "lanes":
                            BcgndShp.get_Cells("Prop.LanesCount").FormulaU = 
                                Td.Attributes["v"].InnerText;
                            break;
                        default:
                            //Если такой ключ не соответсвует данным улицы, добавляем его
                            BcgndShp.AddRow(243, i, 0);
                            ShpCell = BcgndShp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            ShpCell.RowNameU =
                                Td.Attributes["k"].InnerText.Replace(":", "_");
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = BcgndShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                            i++;
                            break;
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
                //throw;
            }
        }
        /// <summary>
        /// Функция заменяет нарисованную из OSM фигуру на фигуру здания
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="shp">Фигура отрисованная из OSM</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateCorrectBuilding(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            ref Visio.Shape shp, System.Xml.XmlNodeList TdList)
        {
            try
            {
                Visio.Master Mstr;
                Visio.Shape BldngShp;
                Visio.Cell ShpCell;
                short i;

                //создаем покрытие (лицевую часть)
                Mstr = VisioApp.Documents["План на местности.vss"].Masters["Здание"];
                BldngShp = VisioApp.ActivePage.Drop(Mstr.Shapes[1], 0, 0);
                i = BldngShp.get_RowCount(243);  //Определяем количество строк в секции visCustomProps

                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Width");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Height");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Angle");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "PinX");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "PinY");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "LocPinX");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "LocPinY");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Geometry1.X1");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Geometry1.Y1");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Geometry1.X2");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Geometry1.Y2");
                DrawTools.CopyCellFormula(ref shp, ref BldngShp, "Geometry1.A2");

                //Указываем свойства здания
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    switch (Td.Attributes["k"].InnerText)
                    {
                        case "building":
                            //Зарезервировано
                            //BldngShp.get_Cells("Prop.addr_housenumber").FormulaU =
                            //    StringToFormulaForString(Td.Attributes["v"].InnerText);
                            break;                        
                        case "addr:housenumber":
                            BldngShp.get_Cells("Prop.addr_housenumber").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                            break;
                        case "addr:street":
                            BldngShp.get_Cells("Prop.addr_street").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);                            
                            break;
                        case "building:levels":
                            BldngShp.get_Cells("Prop.building_levels").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        case "amenity":
                            switch (Td.Attributes["v"].InnerText)
	                        {
                                case "theater":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Театр");
                                    break;
                                case "college":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Колледж");
                                    break;
                                case "kindergarten":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Детский сад");
                                    break;
                                case "library":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Библиотека");
                                    break;
                                case "school":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Школа");
                                    break;
                                case "university":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Университет");
                                    break;
                                case "clinic":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Больница");
                                    break;
                                case "nursing_home":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Дом инвалидов");
                                    break;
                                case "pharmacy":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Аптека");
                                    break;
                                case "arts_centre":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Дом культуры");
                                    break;
                                case "cinema":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Кинотеатр");
                                    break;
                                case "nightclub":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Ночной клуб");
                                    break;
                                case "planetarium":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Планетарий");
                                    break;
                                case "studio":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Телестудия");
                                    break;
                                case "embassy":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Посольство");
                                    break;
                                case "fire_station":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Пожарная часть");
                                    break;
                                case "marketplace":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Рынок");
                                    break;
                                case "place_of_worship":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Культовое учреждение");
                                    break;
                                case "police":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Полиция");
                                    break;
                                case "post_office":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Почта");
                                    break;
                                case "prison":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Тюрьма");
                                    break;
                                case "townhall":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Администрация");
                                    break;
                                case "bank":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString("Банк");
                                    break;
                                default:
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText); 
                                    break;
	                        }
                            break;
                        case "name":
                            BldngShp.get_Cells("Prop.name").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        case "official_name":
                            BldngShp.get_Cells("Prop.name").FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        default:
                            //Если такой тэг не известен - добавляем произвольное свойство
                            BldngShp.AddRow(243, i, 0);
                            ShpCell = BldngShp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            ShpCell.RowNameU =
                                Td.Attributes["k"].InnerText.Replace(":", "_");
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = BldngShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                            
                            i++;
                            break;
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                return false;
                //throw;
            }
        }
        /// <summary>
        /// Прока создает корректную зону Территория
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="shp">Фигура отрисованная из OSM</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateCorrectLandUse(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            ref Visio.Shape shp, System.Xml.XmlNodeList TdList)
        {
            try
            {
                //Visio.Cell ShpCell;
                //short i;

                //Получаем все свойства для данного объекта
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    switch (Td.Attributes["k"].InnerText)
                    {
                        case "landuse":
                            //определяем что за дорога и в соответствии с этим указываем ее ширину
                            switch (Td.Attributes["v"].InnerText)
                            {
                                case "residential":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(235, 241, 222)";
                                    break;
                                case "commercial":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(225,205,203)";
                                    break;
                                case "basin":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(114, 159, 220)";
                                    break;
                                case "construction":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(110, 129, 220)";
                                    break;
                                case "forest":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(176, 221, 128)";
                                    shp.get_Cells("FillBkgnd").FormulaU = "RGB(146, 208, 80)";
                                    shp.get_Cells("FillPattern").FormulaU = "12";
                                    break;
                                case "garages":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(242, 242, 242)";
                                    break;
                                case "grass":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(217, 228, 192)";
                                    shp.get_Cells("FillBkgnd").FormulaU = "RGB(202, 218, 169)";
                                    shp.get_Cells("FillPattern").FormulaU = "11";
                                    break;
                                case "industrial":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(235, 241, 222)";
                                    break;
                                case "landfill":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(166, 166, 166)";
                                    break;
                                case "port":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(235, 241, 222)";
                                    break;
                                case "railway":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(235, 241, 222)";
                                    break;
                                case "village_green":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(176, 221, 128)";
                                    shp.get_Cells("FillBkgnd").FormulaU = "RGB(146, 208, 80)";
                                    shp.get_Cells("FillPattern").FormulaU = "2";                                    
                                    break;
                                default:
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(202,218,169)";
                                    break;
                            }
                            break;
                        default:
                            //Если такой ключ не соответсвует данным улицы, добавляем его
                            //LndUseShp.AddRow(243, i, 0);
                            //ShpCell = LndUseShp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            //ShpCell.RowNameU =
                            //    Td.Attributes["k"].InnerText.Replace(":", "_");
                            //ShpCell.FormulaU =
                            //    StringToFormulaForString(Td.Attributes["k"].InnerText);
                            //ShpCell = LndUseShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            //ShpCell.FormulaU =
                            //    StringToFormulaForString(Td.Attributes["v"].InnerText);
                            //i++;
                            break;
                    }
                }
                shp.get_CellsSRC(1, 6, 0).FormulaU = GetLayerNumber(ref VisioApp, "Территории");  //visSectionObject = 1, visRowLayerMem = 6, visLayerMember = 0
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("Территории: " + err.Message);
                return false;
                //throw;
            }
        }

        /// <summary>
        /// Прока создает корректную зону Зоны отдыха
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="shp">Фигура отрисованная из OSM</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateCorrectLeisure(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            ref Visio.Shape shp, System.Xml.XmlNodeList TdList)
        {
            try
            {
                //Visio.Cell ShpCell;
                //short i;

                //Получаем все свойства для данного объекта
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    switch (Td.Attributes["k"].InnerText)
                    {
                        case "leisure":
                            //определяем что за дорога и в соответствии с этим указываем ее ширину
                            switch (Td.Attributes["v"].InnerText)
                            {
                                case "common":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(146, 208, 80)";
                                    break;
                                case "playground":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(209, 235, 241)";
                                    break;
                                case "stadium":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(0, 176, 80)";
                                    break;
                                case "sports_centre":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(0, 176, 80)";
                                    break;
                                case "track":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(202, 218, 169)";
                                    break;
                                case "park":
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(146, 208, 80)";
                                    break;  
                                default:
                                    shp.get_Cells("FillForegnd").FormulaU = "RGB(146, 208, 80)";
                                    break;
                            }
                            break;
                        default:
                            //Если такой ключ не соответсвует данным улицы, добавляем его
                            //LndUseShp.AddRow(243, i, 0);
                            //ShpCell = LndUseShp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            //ShpCell.RowNameU =
                            //    Td.Attributes["k"].InnerText.Replace(":", "_");
                            //ShpCell.FormulaU =
                            //    StringToFormulaForString(Td.Attributes["k"].InnerText);
                            //ShpCell = LndUseShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            //ShpCell.FormulaU =
                            //    StringToFormulaForString(Td.Attributes["v"].InnerText);
                            //i++;
                            break;
                    }
                }
                shp.get_CellsSRC(1, 6, 0).FormulaU = GetLayerNumber(ref VisioApp, "Зоны отдыха");  //visSectionObject = 1, visRowLayerMem = 6, visLayerMember = 0
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("Зоны отдыха: " + err.Message);
                return false;
                //throw;
            }
        }
        /// <summary>
        /// Прока создает корректную зону Ограждение
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="shp">Фигура отрисованная из OSM</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateCorrectBorder(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            ref Visio.Shape shp, System.Xml.XmlNodeList TdList)
        {
            try
            {
                Visio.Cell ShpCell;
                short i=0;

                //Получаем все свойства для данного объекта
                foreach (System.Xml.XmlNode Td in TdList)
                {
                    switch (Td.Attributes["k"].InnerText)
                    {
                        case "barrier":
                            //определяем что за дорога и в соответствии с этим указываем ее ширину
                            switch (Td.Attributes["v"].InnerText)
                            {
                                case "fence":
                                    shp.get_CellsSRC(10, 0, 0).FormulaU = "1";
                                    break;
                                default:
                                    shp.get_CellsSRC(10, 0, 0).FormulaU = "1";                                    
                                    break;
                            }
                            break;
                        default:
                            //Если такой ключ не соответсвует данным улицы, добавляем его
                            shp.AddRow(243, i, 0);
                            ShpCell = shp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            ShpCell.RowNameU =
                                Td.Attributes["k"].InnerText.Replace(":", "_");
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = shp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                DrawTools.StringToFormulaForString(Td.Attributes["v"].InnerText);
                            i++;
                            break;
                    }
                }
                shp.get_CellsSRC(1, 6, 0).FormulaU = GetLayerNumber(ref VisioApp, "Зоны отдыха");  //visSectionObject = 1, visRowLayerMem = 6, visLayerMember = 0
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("Ограждения: " + err.Message);
                return false;
                //throw;
            }
        }


        /// <summary>
        /// Прока фиксит расположение фигур определнных слоев
        /// Здания и ограды (сначала ограды, потом здания) - вперед, Территории назад
        /// </summary>
        /// <param name="VisioApp">Активное приложение Visio</param>
        private void LayersFix(ref Microsoft.Office.Interop.Visio.Application VisioApp)
        {
            Visio.Selection LayerSelection;
            Visio.Page CurPage;

            try
            {
                //Выбираем текущую страницу
                CurPage =  VisioApp.ActiveWindow.Page;
                
                //Очищаем имеющиеся выделения
                LayerSelection = VisioApp.ActiveWindow.Selection;
                LayerSelection.DeselectAll();

                //Отправляем назад фигуры Зон отдыха
                LayerSelection = CurPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper, "Зоны отдыха");
                LayerSelection.SendToBack();
                LayerSelection.DeselectAll();

                //Отправляем назад фигуры территории
                LayerSelection = CurPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper, "Территории");
                LayerSelection.SendToBack();
                LayerSelection.DeselectAll();

                //Отправляем вперед фигуры зданий
                LayerSelection = CurPage.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper, "Здания");
                LayerSelection.BringToFront();
                LayerSelection.DeselectAll();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                //throw;
            }
        }

        #endregion Проки и функции отрисовки


        //------------------------проки и функции масштаба---------------------------
        /// <summary>
        /// Прока задает размер и масштаб листа в зависимости от размера листа
        /// </summary>
        /// <param name="Page">Старница</param>
        /// <param name="a_Box">Координат прямоугольника</param>
        private void SetSizeScale(ref Microsoft.Office.Interop.Visio.Application VisioApp, DrawTools.CoordRecatangle a_Box)
        {
            Visio.Page v_Page;
            
            Double LenightHor = 0;
            Double LenightVert = 0;
            try
            {
                //Получаем линейные размеры прямоугольника
                DrawTools.GetSizes(a_Box, ref LenightHor, ref LenightVert);
                //MessageBox.Show(LenightHor.ToString());

                //Устанавливаем масштаб в зависмости от размеров прямоугольинка
                v_Page = VisioApp.ActivePage;
                v_Page.PageSheet.get_Cells("DrawingScaleType").Formula = "3";
                v_Page.PageSheet.get_Cells("PageWidth").Formula = LenightHor + " m";
                v_Page.PageSheet.get_Cells("PageHeight").Formula = LenightVert + " m";                
                
                if (LenightHor < (Double)200) //масштаб 1:200
                {
                    v_Page.PageSheet.get_Cells("PageScale").Formula = "10 mm";
                    v_Page.PageSheet.get_Cells("DrawingScale").Formula = "2 m";
                }
                else if (LenightHor >= (Double)200 && LenightHor < (Double)500) //масштаб 1:500
                {
                    v_Page.PageSheet.get_Cells("PageScale").Formula = "10 mm";
                    v_Page.PageSheet.get_Cells("DrawingScale").Formula = "5 m";
                }
                else if (LenightHor >= (Double)500 && LenightHor < (Double)1000) //масштаб 1:1000
                {
                    v_Page.PageSheet.get_Cells("PageScale").Formula = "10 mm";
                    v_Page.PageSheet.get_Cells("DrawingScale").Formula = "10 m";
                }
                else if (LenightHor >= (Double)1000 && LenightHor < (Double)5000) //масштаб 1:5000
                {
                    v_Page.PageSheet.get_Cells("PageScale").Formula = "10 mm";
                    v_Page.PageSheet.get_Cells("DrawingScale").Formula = "50 m";
                }                
                else //масштаб 1:10000
                {
                    v_Page.PageSheet.get_Cells("PageScale").Formula = "10 mm";
                    v_Page.PageSheet.get_Cells("DrawingScale").Formula = "100 m";
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                throw;
            }
        }


        //------------------------Служебные функции----------------------------------
        private void SetBuildingData(ref Visio.Shape ShapeFrom, ref Visio.Shape ShapeTo, 
            String CellNemeFrom, String CellNemeTo)
        {
            try
            {
                ShapeTo.get_Cells(CellNemeTo).FormulaU =
                    DrawTools.StringToFormulaForString(ShapeFrom.get_Cells(CellNemeFrom).get_ResultStr("visUnitsString"));
            }
            catch (Exception)
            {
                //MessageBox.Show(StringToFormulaForString(ShapeFrom.get_Cells(CellNemeFrom).get_ResultStr("visUnitsString")));
                //throw;
            }
        }

        private string GetLayerNumber(ref Microsoft.Office.Interop.Visio.Application VisioApp, String LayerName)
        {
            String ResultLayerString = "_";
            Visio.Page CurPage;
            int LayerNumber;

            try
            {
                CurPage = VisioApp.ActiveWindow.Page;

                for (int i = 1; i < CurPage.Layers.Count + 1; i++)
                {
                    if (CurPage.Layers[i].Name == LayerName)
                    {
                        LayerNumber = i - 1;
                        ResultLayerString = DrawTools.StringToFormulaForString(LayerNumber.ToString());
                        return ResultLayerString;
                    }
                }
                CurPage.Layers.Add(LayerName);
                LayerNumber = CurPage.Layers.Count - 1;
                ResultLayerString = DrawTools.StringToFormulaForString(LayerNumber.ToString());

                return ResultLayerString;
            }
            catch (Exception err)
            {
                MessageBox.Show("Слой: " + err.Message);
                return "";
                //throw;
            }
        }

        static private System.Xml.XmlDocument GetKML(string _filepath, string _kmlfile = "doc.KML")
        {
            System.Xml.XmlDocument tempXMLDoc = new System.Xml.XmlDocument();

            var zip = ZipFile.Read(_filepath);

            //Ionic.Crc.CrcCalculatorStream kmzreader = zip[_kmlfile].OpenReader();
            //StreamReader kmlreader = new StreamReader(kmzreader);
            StreamReader kmlreader = new StreamReader(zip[_kmlfile].OpenReader());
            tempXMLDoc.LoadXml(kmlreader.ReadToEnd());
            kmlreader.Close();

            return tempXMLDoc;
        }



        //-----------------Публичные проки измененеия состояния прогрессбара--------------------
        public void SetProgressbarMaximum(int maxVal)
        {
            this.PrB_DrawProcess.Maximum = maxVal;
        }
        public void SetProgressBarCurrentValue(int curVal)
        {
            this.PrB_DrawProcess.Value = curVal;
        }

    }
}