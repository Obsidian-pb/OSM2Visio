using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Xml;

namespace OSM2Visio
{
    public partial class f_ImportDataDialog : Form
    {
        private Visio.Application VisApp;
        
        public f_ImportDataDialog()
        {
            InitializeComponent();
        }

        private void B_Search_Click(object sender, EventArgs e)
        {
            FD.Filter = "Файл данных OSM|*.osm";
            FD.ShowDialog();
            TB_FilePath.Text = FD.FileName;
        }

        private void B_SearchEWS_Click(object sender, EventArgs e)
        {
            //Указываем, путь к источнику данных
            switch (CB_EWSSource.SelectedIndex)
            {
                case 0:
                    FD.Filter = "Файл данных OSM|*.osm";
                    break;
                case 1:
                    FD.Filter = "Файл БД EWS|*.mdb";
                    break;
                case 2:
                    FD.Filter = "Файл строки подключения к БД|*.txt";
                    break;
                case 3:
                    FD.Filter = "Файл данных ЭСУ ППВ|*.kmz";
                    break;
                default:
                    FD.Filter = "Все файлы|*.*";
                    break;
            }
            FD.ShowDialog();  //открываем диалог
            TB_EWSPath.Text = FD.FileName;
        }

        private void CB_EWSSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CB_EWSSource.SelectedIndex == 0)
            {
                TB_EWSPath.Enabled = false;
                B_SearchEWS.Enabled = false;
            }
            else
            {
                TB_EWSPath.Enabled = true;
                B_SearchEWS.Enabled = true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Открывается окно со справкой по OSM");
        }

        private void B_OK_Click(object sender, EventArgs e)
        {
            System.Xml.XmlDocument OSMData = new System.Xml.XmlDocument();
            
            //Получаем ссылку на текущее приложение
            Visio.Application VisApp = Globals.ThisAddIn.Application;

            //Формируем документ XML
            OSMData.Load(TB_FilePath.Text);
            
            //Получаем путь к файлу данных ИНППВ
            string EWS_DataFilePath = TB_EWSPath.Text;

            //Закрываем текущую форму
            this.Close();

            //Создаем экземпляр формы процесса отрисовки
            f_DrawProcess v_ProcessForm = new f_DrawProcess();

            v_ProcessForm.Pv_Draw(VisApp, OSMData, CB_EWSSource.SelectedIndex, EWS_DataFilePath);
            //v_ProcessForm.Show();
        }
    }
}
