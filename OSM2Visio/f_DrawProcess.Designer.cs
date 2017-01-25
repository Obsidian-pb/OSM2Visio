using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
//using System.Array;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

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

        private double EARTH_RADIUS = 6371032;
        private double PI = 3.141592654;
        private double INCHINMETER = 39.3701;

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


        //------------Проки отрисовки
        //---------------------------Внешние проки формы
        //Основная прока отрисовки зданий из OSM, Получает XMLDocument документ с данными из файла
        public void Pv_Draw(Microsoft.Office.Interop.Visio.Application VisioApp,
            System.Xml.XmlDocument Data)
        {
            //Переменные для работы
            System.Xml.XmlNodeList NodesList;
            System.Xml.XmlNodeList NdList;
            System.Xml.XmlNodeList TdList;

            CoordRecatangle v_Box = new CoordRecatangle();

            //Dim Crdnt As Coordinate  '- доработать
            Double x = 0;
            Double y = 0;
            Double XPos;
            Double YPos;

            Double InchInGradH;  //Количество дюймов в одном градусе долготы
            Double InchInGradV;  //Количество дюймов в одном градусе широты

            Visio.Shape shp;

            //Visio.Cell ShpCell;
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

            pb_GetBoundBox(Data, ref v_Box);

            //---Определяем количество дюймов в одном градусе долготы
            InchInGradH = GetInchesInGradH(v_Box);
            InchInGradV = GetInchesInGradV(v_Box);
            //MessageBox.Show("Размеры вычислены: Гор -> " + InchInGradH.ToString("R"));

            //---Определяем линейные размеры прямоугольника и приравниваем к нему рабочий лист
            SetSizeScale(ref VisioApp, v_Box);
            //---Увеличиваме картинку листа по размеру окна  Application.ActiveWindow.ViewFit = visFitPage
            VisioApp.ActiveWindow.ViewFit = (int)Visio.VisWindowFit.visFitPage;
            this.Focus();
            this.Top = 200; this.Left = 400;

            //---Получаем узел с перечислением Way
            NodesList = Data.SelectNodes("//way");
            //---Указываем максимальное значение процессбара
            this.Text = "Рисуются полигоны";
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
                    
                    GetPosition(Nd.Attributes["ref"].InnerText, ref Data, ref x, ref y);
                    
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
                        PolyLineToLine(ref shp);
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
            
            //---Получаем список узлов с перечислением node
            NodesList = Data.SelectNodes("//node");
            //---Указываем максимальное значение процессбара
            this.Text = "Расставляются водоисточники";
            this.PrB_DrawProcess.Maximum = NodesList.Count;
            i = 0;

            Coordinate pnt; pnt.x = 0; pnt.y = 0;

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
                            GetPosition(node.Attributes["id"].InnerText, ref Data, ref x, ref y);

                            //Получаем координату относительно края области (в дюймах - все в дюймах)
                            XPos = (x - v_Box.XY1.x);   YPos = (y - v_Box.XY1.y);
                            pnt.x = XPos * InchInGradH; pnt.y = YPos * InchInGradV;
                           
                            //Создаем новый ИНППВ, согласно указанным в node координатам
                            CreateEWS_OSM(ref VisioApp, TdList, Td.Attributes["v"].InnerText, pnt);
                            //NeedDelete = true;
                        }
                    }

                }

                this.PrB_DrawProcess.Value = i;
                i++;
            }

            //Распределяем слои - здания вперед, территории назад
            LayersFix(ref VisioApp);

            //Отчет о завршении
            MessageBox.Show("Отрисовано " + i.ToString() + " объектов");
            this.B_OK.Enabled = true;

        }



        //------------------------Проки и функции отрисовки
        //Прока передает в переменные данные об относительном положении точки на листе
        private Boolean GetPosition(String NodeID, ref System.Xml.XmlDocument NodeDoc,
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

                CopyCellFormula(ref shp, ref BcgndShp, "Width");
                CopyCellFormula(ref shp, ref BcgndShp, "Height");
                CopyCellFormula(ref shp, ref BcgndShp, "Angle");
                CopyCellFormula(ref shp, ref BcgndShp, "PinX");
                CopyCellFormula(ref shp, ref BcgndShp, "PinY");
                CopyCellFormula(ref shp, ref BcgndShp, "LocPinX");
                CopyCellFormula(ref shp, ref BcgndShp, "LocPinY");
                CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.X1");
                CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.Y1");
                CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.X2");
                CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.Y2");
                CopyCellFormula(ref shp, ref BcgndShp, "Geometry1.A2");

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
                                StringToFormulaForString(Td.Attributes["v"].InnerText);
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
                                StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = BcgndShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);
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

                CopyCellFormula(ref shp, ref BldngShp, "Width");
                CopyCellFormula(ref shp, ref BldngShp, "Height");
                CopyCellFormula(ref shp, ref BldngShp, "Angle");
                CopyCellFormula(ref shp, ref BldngShp, "PinX");
                CopyCellFormula(ref shp, ref BldngShp, "PinY");
                CopyCellFormula(ref shp, ref BldngShp, "LocPinX");
                CopyCellFormula(ref shp, ref BldngShp, "LocPinY");
                CopyCellFormula(ref shp, ref BldngShp, "Geometry1.X1");
                CopyCellFormula(ref shp, ref BldngShp, "Geometry1.Y1");
                CopyCellFormula(ref shp, ref BldngShp, "Geometry1.X2");
                CopyCellFormula(ref shp, ref BldngShp, "Geometry1.Y2");
                CopyCellFormula(ref shp, ref BldngShp, "Geometry1.A2");

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
                                StringToFormulaForString(Td.Attributes["v"].InnerText);
                            break;
                        case "addr:street":
                            BldngShp.get_Cells("Prop.addr_street").FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);                            
                            break;
                        case "building:levels":
                            BldngShp.get_Cells("Prop.building_levels").FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        case "amenity":
                            switch (Td.Attributes["v"].InnerText)
	                        {
                                case "theater":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Театр");
                                    break;
                                case "college":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Колледж");
                                    break;
                                case "kindergarten":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Детский сад");
                                    break;
                                case "library":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Библиотека");
                                    break;
                                case "school":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Школа");
                                    break;
                                case "university":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Университет");
                                    break;
                                case "clinic":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Больница");
                                    break;
                                case "nursing_home":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Дом инвалидов");
                                    break;
                                case "pharmacy":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Аптека");
                                    break;
                                case "arts_centre":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Дом культуры");
                                    break;
                                case "cinema":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Кинотеатр");
                                    break;
                                case "nightclub":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Ночной клуб");
                                    break;
                                case "planetarium":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Планетарий");
                                    break;
                                case "studio":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Телестудия");
                                    break;
                                case "embassy":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Посольство");
                                    break;
                                case "fire_station":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Пожарная часть");
                                    break;
                                case "marketplace":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Рынок");
                                    break;
                                case "place_of_worship":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Культовое учреждение");
                                    break;
                                case "police":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Полиция");
                                    break;
                                case "post_office":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Почта");
                                    break;
                                case "prison":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Тюрьма");
                                    break;
                                case "townhall":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Администрация");
                                    break;
                                case "bank":
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString("Банк");
                                    break;
                                default:
                                    BldngShp.get_Cells("Prop.amenity").FormulaU =
                                        StringToFormulaForString(Td.Attributes["v"].InnerText); 
                                    break;
	                        }
                            break;
                        case "name":
                            BldngShp.get_Cells("Prop.name").FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        case "official_name":
                            BldngShp.get_Cells("Prop.name").FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);                              
                            break;
                        default:
                            //Если такой тэг не известен - добавляем произвольное свойство
                            BldngShp.AddRow(243, i, 0);
                            ShpCell = BldngShp.get_CellsSRC(243, i, 2);  //visCustPropsLabel
                            ShpCell.RowNameU =
                                Td.Attributes["k"].InnerText.Replace(":", "_");
                            ShpCell.FormulaU =
                                StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = BldngShp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);
                            
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
                                StringToFormulaForString(Td.Attributes["k"].InnerText);
                            ShpCell = shp.get_CellsSRC(243, i, 0);  //visCustPropsValue
                            ShpCell.FormulaU =
                                StringToFormulaForString(Td.Attributes["v"].InnerText);
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




        //-------------------------------------------Вставка УГО ИНППВ--------------------------------------

        /// <summary>
        /// Прока вставляет значек ИНППВ
        /// </summary>
        /// <param name="VisioApp">Текущее приложение Visio</param>
        /// <param name="TdList">перечень узлов с данными (tag)</param>
        /// <returns></returns>
        private Boolean CreateEWS_OSM(ref Microsoft.Office.Interop.Visio.Application VisioApp,
            System.Xml.XmlNodeList TdList, String EWSType, Coordinate pnt)
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
                    shp.get_Cells("Prop.PGAdress").FormulaU = StringToFormulaForString(" ");
                    shp.get_Cells("Prop.PGNumber").FormulaU = StringToFormulaForString(" ");

                    //Получаем данные о ПГ
                    foreach (System.Xml.XmlNode Td in TdList)
                    {
                        switch (Td.Attributes["k"].InnerText)
                        {
                            case "fire_hydrant:street":
                                shp.get_Cells("Prop.PGAdress").FormulaU = StringToFormulaForString(Td.Attributes["v"].InnerText + " " + shp.get_Cells("Prop.PGAdress").FormulaU);
                                break;
                            case "fire_hydrant:housenumber":
                                shp.get_Cells("Prop.PGAdress").FormulaU = StringToFormulaForString(shp.get_Cells("Prop.PGAdress").FormulaU + " " + Td.Attributes["v"].InnerText);
                                break;
                            case "ref":
                                shp.get_Cells("Prop.PGNumber").FormulaU = StringToFormulaForString(Td.Attributes["v"].InnerText);
                                break;
                            case "fire_hydrant:diameter":
                                tmpStr = Td.Attributes["v"].InnerText;
                                if (tmpStr.IndexOf("К")>0)
                                    shp.get_Cells("Prop.PipeType").FormulaU = "INDEX(0, Prop.PipeType.Format)";
                                if (tmpStr.IndexOf("Т") > 0)
                                    shp.get_Cells("Prop.PipeType").FormulaU = "INDEX(1, Prop.PipeType.Format)";

                                //Убираем из строки ненужные символы
                                tmpStr = tmpStr.Replace("К", ""); tmpStr = tmpStr.Replace("Т", ""); tmpStr = tmpStr.Replace("-", "");

                                shp.get_Cells("Prop.PipeDiameter").FormulaU = StringToFormulaForString(tmpStr);
                                break;
                            case "fire_hydrant:pressure":
                                tmpStr = Td.Attributes["v"].InnerText;
                                switch (tmpStr)
                                {
                                    case "1":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("10");
                                        break;
                                    case "2":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("20");
                                        break;
                                    case "3":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("30");
                                        break;
                                    case "4":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("40");
                                        break;
                                    case "5":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("50");
                                        break;
                                    case "6":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("60");
                                        break;
                                    case "7":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("70");
                                        break;
                                    case "8":
                                        shp.get_Cells("Prop.Pressure").FormulaU = StringToFormulaForString("80");
                                        break;
                                    case "no":
                                        shp.get_Cells("LineColor").FormulaU = StringToFormulaForString("2");
                                        shp.get_Cells("Char.Color").FormulaU = StringToFormulaForString("2");
                                        break;
                                }
                                break;
                            default:
                                break;
                        }
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
                                shp.get_Cells("Prop.PWAdress").FormulaU = StringToFormulaForString(Td.Attributes["v"].InnerText + " " + shp.get_Cells("Prop.PWAdress").FormulaU);
                                break;
                            case "water_tank:housenumber":
                                shp.get_Cells("Prop.PWAdress").FormulaU = StringToFormulaForString(shp.get_Cells("Prop.PWAdress").FormulaU + " " + Td.Attributes["v"].InnerText);
                                break;
                            case "ref":
                                shp.get_Cells("Prop.PWNumber").FormulaU = StringToFormulaForString(Td.Attributes["v"].InnerText);
                                break;
                            case "water_tank:volume":
                                tmpStr = Td.Attributes["v"].InnerText;
                                if (tmpStr.IndexOf("no") > 0)
                                {
                                    shp.get_Cells("LineColor").FormulaU = StringToFormulaForString("2");
                                    shp.get_Cells("Char.Color").FormulaU = StringToFormulaForString("2");
                                    tmpStr = "0";
                                }

                                shp.get_Cells("Prop.PWValue").FormulaU = StringToFormulaForString(tmpStr);
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




        //------------------------проки и функции масштаба---------------------------
        /// <summary>
        /// Прока задает размер и масштаб листа в зависимости от размера листа
        /// </summary>
        /// <param name="Page">Старница</param>
        /// <param name="a_Box">Координат прямоугольника</param>
        private void SetSizeScale(ref Microsoft.Office.Interop.Visio.Application VisioApp, CoordRecatangle a_Box)
        {
            Visio.Page v_Page;
            
            Double LenightHor = 0;
            Double LenightVert = 0;
            try
            {
                //Получаем линейные размеры прямоугольника
                GetSizes(a_Box, ref LenightHor, ref LenightVert);
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
        /// <summary>
        /// Прока возвраает линейные размеры прямоугольника
        /// </summary>
        /// <param name="a_Box">Прямоугольник</param>
        /// <param name="HorLen">Горизонатльная длина (по параллели)</param>
        /// <param name="VertLen">Вертикальная длина (по меридиану)</param>
        private Boolean GetSizes(CoordRecatangle a_Box, ref Double HorLen, ref Double VertLen)
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
                MessageBox.Show(err.Message);
                return false;
                throw;
            }
        }

        //------------------------Служебные функции----------------------------------
        //Функция возварщает количество дюймов в одном градусе долготы
        private Double GetInchesInGradH(CoordRecatangle a_Box)
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
        private Double GetInchesInGradV(CoordRecatangle a_Box)
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

        private Boolean pb_GetBoundBox(System.Xml.XmlDocument Data, ref CoordRecatangle a_Box)
        {
            //Функция получает DOM документ с данными из OSM и рамку с координатами,
            //в которую соохарняет данные о границе выборки из OSM. Возвращает True, если функция отработала и False если нет
            System.Xml.XmlNode vCR_BoundsNode;
            try
            {
                vCR_BoundsNode = Data.SelectNodes("//bounds")[0];
                a_Box.XY1.SetCrdnt(pf_StrToDbl(vCR_BoundsNode.Attributes["minlon"].InnerText),
                        pf_StrToDbl(vCR_BoundsNode.Attributes["minlat"].InnerText));
                a_Box.XY2.SetCrdnt(pf_StrToDbl(vCR_BoundsNode.Attributes["maxlon"].InnerText),
                        pf_StrToDbl(vCR_BoundsNode.Attributes["maxlat"].InnerText));
                return true;
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        }

        //Функция корректно преобразует строку к типу Double
        private Double pf_StrToDbl(String a_Str)
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
        /// Прока превращает входящую строку в строку формулы для вставки 
        /// в ячейки Visio. Заменяем все двойные кавычки(") парой двойных кавычек("")
        /// в начале и в конце строки.</summary>
        /// <param name="inputValue">Входящая строка</param>
        /// <returns>измененная строка, которая может быть программно
        /// назначена ячейке. Не может быть напрямую назначена ячейке, поскольку не имеет
        /// "=" в начале.</returns>
        public string StringToFormulaForString(string inputValue)
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
        /// Прока связывает ячейку двух фигур 
        /// </summary>
        /// <param name="a_shpFace"></param>
        /// <param name="a_shpBckgnd"></param>
        /// <param name="a_CellName"></param>
        private void LinkShapes(ref Visio.Shape a_shpFace, ref Visio.Shape a_shpBckgnd, string a_CellName)
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
        private void CopyCellFormula(ref Visio.Shape a_shpOrigin, ref Visio.Shape a_shpDescendent, string a_CellName)
        {
            try
            {
                //Visio.Cell ShpCell;
                //Visio.Cell ShpCell;

                //ShpCell = a_shpBckgnd.get_Cells(a_CellName);
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

        private void SetBuildingData(ref Visio.Shape ShapeFrom, ref Visio.Shape ShapeTo, 
            String CellNemeFrom, String CellNemeTo)
        {
            try
            {
                ShapeTo.get_Cells(CellNemeTo).FormulaU =
                    StringToFormulaForString(ShapeFrom.get_Cells(CellNemeFrom).get_ResultStr("visUnitsString"));
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
                        ResultLayerString = StringToFormulaForString(LayerNumber.ToString());
                        return ResultLayerString;
                    }
                }
                CurPage.Layers.Add(LayerName);
                LayerNumber = CurPage.Layers.Count - 1;
                ResultLayerString = StringToFormulaForString(LayerNumber.ToString());

                return ResultLayerString;
            }
            catch (Exception err)
            {
                MessageBox.Show("Слой: " + err.Message);
                return "";
                //throw;
            }
        }

        /// <summary>
        /// Прока обращает фигуру нарисованную при помощи метода PolylineTo
        /// в фигуру нарисованную при помощи метода LineTo
        /// </summary>
        /// <param name="VisioApp">Активное приложение Visio</param>
        /// <param name="shp">Фигура</param>
        private void PolyLineToLine(ref Visio.Shape shp)
        {
            string PolyLineString;
            int i;
            short rowIndex=2;

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
                        shp.get_CellsSRC(10, rowIndex, 0).Formula = "Width*" + GetStrByIndex(PolyLineString, i-1).ToString();
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

        private Double GetStrByIndex(string str, int index)
        {
            int i;
            int count = 0;
            int pos1 = 0;
            int pos2 = str.Length;
            string DblStr;
            Double DblVal;


            if (GetItemsCount(str) == 0)
                return 0; //pf_StrToDbl(str);
            //MessageBox.Show("1 -> " + str);
            try
            {
                for (i = 0; i < str.Length; i++)
                {
                    if (str.Substring(i, 1) == ";")
                    {
                        pos2 = i;
                        count++;
                        //MessageBox.Show(i.ToString());
                        if (count == index)
                        {
                            DblStr = str.Substring(pos1, pos2 - pos1);
                            //MessageBox.Show(str + " - > " + i.ToString() + ", index=" + index.ToString() + " text=" + DblStr);
                            return pf_StrToDbl(DblStr);
                            //return Convert.ToDouble(str.Substring(pos1, pos2 - pos1));
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

        private int GetItemsCount(string str)
        {
            int i;
            int count = 0;

            for (i = 0; i < str.Length; i++)
            {
                if (str.Substring(i, 1) == ";")
                {
                    count++;
                    //MessageBox.Show(count.ToString());
                }
            }
            return count;
        }
    }
}