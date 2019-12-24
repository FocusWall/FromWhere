using System;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Components;
using MetroFramework.Forms;

namespace Mail
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            AutoCompleteStringCollection source = new AutoCompleteStringCollection()
            {
            "Кузнецов Валерий Семенович",
            "Иванов",
            "Петров",
            "Кустов"
            };

            textBox1.AutoCompleteCustomSource = source;
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;

            AutoCompleteStringCollection source2 = new AutoCompleteStringCollection()
            {
            "г. Новосибирск, ул. Кирова 63",
            "Иванов2",
            "Петров2",
            "Кустов2"
            };

            textBox2.AutoCompleteCustomSource = source2;
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;

            AutoCompleteStringCollection source3 = new AutoCompleteStringCollection()
            {
            "Шевцову Ярославу Петровичу",
            "Иванов2",
            "Петров2",
            "Кустов3"
            };

            textBox3.AutoCompleteCustomSource = source3;
            textBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox3.AutoCompleteSource = AutoCompleteSource.CustomSource;

            AutoCompleteStringCollection source4 = new AutoCompleteStringCollection()
            {
            "г. Новосибирск, ул. Кирова 63",
            "Иванов2",
            "Петров2",
            "Кустов4"
            };

            textBox4.AutoCompleteCustomSource = source4;
            textBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox4.AutoCompleteSource = AutoCompleteSource.CustomSource;

            AutoCompleteStringCollection source5 = new AutoCompleteStringCollection()
            {
            "860007",
            "Иванов2",
            "Петров2",
            "Кустов5"
            };

            textBox5.AutoCompleteCustomSource = source5;
            textBox5.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox5.AutoCompleteSource = AutoCompleteSource.CustomSource;

            AutoCompleteStringCollection source6 = new AutoCompleteStringCollection()
            {
            "860007",
            "Иванов2",
            "Петров2",
            "Кустов6"
            };

            textBox6.AutoCompleteCustomSource = source6;
            textBox6.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox6.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void пустойToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            String from = textBox1.Text;
            String fromAdress = textBox2.Text;
            String fromIndex = textBox5.Text;
            String to = textBox3.Text;
            String toAdress = textBox4.Text;
            String toIndex = textBox6.Text;
            //Создание страницы (Конверт С6)
            PdfDocument print = new PdfDocument();
            PdfPage page = print.AddPage();
            page.Height = XUnit.FromMillimeter(114);
            page.Width = XUnit.FromMillimeter(162);
            XGraphics gfx = XGraphics.FromPdfPage(page);
            //Стили для документа
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFont font = new XFont("Times New Roman", 14, XFontStyle.Italic, options);
            XFont font2 = new XFont("Times New Roman", 14, XFontStyle.Underline, options);
            XFont fontIndex = new XFont("Times New Roman", 8, XFontStyle.Italic, options);
            XFont fontIndex2 = new XFont("Times New Roman", 10, XFontStyle.Bold, options);

            XPen line = new XPen(XColors.Black, 1);
            XPen lineForIndex = new XPen(XColors.Black, 5);
            XPen lineForIndex2 = new XPen(XColors.Black, 2);
            XPen DashLine = new XPen(XColors.Black, 1);
            DashLine.DashStyle = XDashStyle.Dash;

            // Отрисовка индекса слева внизу
                // Толстый пунктир
            gfx.DrawLine(lineForIndex, 140, 250, 160, 250);
            gfx.DrawLine(lineForIndex, 115, 250, 135, 250);
            gfx.DrawLine(lineForIndex, 90, 250, 110, 250);
            gfx.DrawLine(lineForIndex, 65, 250, 85, 250);
            gfx.DrawLine(lineForIndex, 40, 250, 60, 250);
            gfx.DrawLine(lineForIndex, 15, 250, 35, 250);
            gfx.DrawLine(lineForIndex, 165, 250, 185, 250);
                // Черточка слева
            gfx.DrawLine(lineForIndex2, 15, 260, 35, 260);
                // Крышечки сверху
            gfx.DrawLine(DashLine, 40, 260, 60, 260);
            gfx.DrawLine(DashLine, 65, 260, 85, 260);
            gfx.DrawLine(DashLine, 90, 260, 110, 260);
            gfx.DrawLine(DashLine, 115, 260, 135, 260);
            gfx.DrawLine(DashLine, 140, 260, 160, 260);
            gfx.DrawLine(DashLine, 165, 260, 185, 260);
                // Диагональные линии
            gfx.DrawLine(DashLine, 40, 280, 60, 260);
            gfx.DrawLine(DashLine, 40, 300, 60, 280);

            gfx.DrawLine(DashLine, 65, 280, 85, 260);
            gfx.DrawLine(DashLine, 65, 300, 85, 280);

            gfx.DrawLine(DashLine, 90, 280, 110, 260);
            gfx.DrawLine(DashLine, 90, 300, 110, 280);

            gfx.DrawLine(DashLine, 115, 280, 135, 260);
            gfx.DrawLine(DashLine, 115, 300, 135, 280);

            gfx.DrawLine(DashLine, 140, 280, 160, 260);
            gfx.DrawLine(DashLine, 140, 300, 160, 280);

            gfx.DrawLine(DashLine, 165, 280, 185, 260);
            gfx.DrawLine(DashLine, 165, 300, 185, 280);
                // Нижние крышечки
            gfx.DrawLine(DashLine, 40, 300, 60, 300);
            gfx.DrawLine(DashLine, 65, 300, 85, 300);
            gfx.DrawLine(DashLine, 90, 300, 110, 300);
            gfx.DrawLine(DashLine, 115, 300, 135, 300);
            gfx.DrawLine(DashLine, 140, 300, 160, 300);
            gfx.DrawLine(DashLine, 165, 300, 185, 300);
                // Средние крышечки
            gfx.DrawLine(DashLine, 40, 280, 60, 280);
            gfx.DrawLine(DashLine, 65, 280, 85, 280);
            gfx.DrawLine(DashLine, 90, 280, 110, 280);
            gfx.DrawLine(DashLine, 115, 280, 135, 280);
            gfx.DrawLine(DashLine, 140, 280, 160, 280);
            gfx.DrawLine(DashLine, 165, 280, 185, 280);
                //Палки слева
            gfx.DrawLine(DashLine, 40, 260, 40, 300);
            gfx.DrawLine(DashLine, 65, 260, 65, 300);
            gfx.DrawLine(DashLine, 90, 260, 90, 300);
            gfx.DrawLine(DashLine, 115, 260, 115, 300);
            gfx.DrawLine(DashLine, 140, 260, 140, 300);
            gfx.DrawLine(DashLine, 165, 260, 165, 300);
                //Палки справа
            gfx.DrawLine(DashLine, 60, 260, 60, 300);
            gfx.DrawLine(DashLine, 85, 260, 85, 300);
            gfx.DrawLine(DashLine, 110, 260, 110, 300);
            gfx.DrawLine(DashLine, 135, 260, 135, 300);
            gfx.DrawLine(DashLine, 160, 260, 160, 300);
            gfx.DrawLine(DashLine, 185, 260, 185, 300);

            // Отрисовка правого верхнего угла
            gfx.DrawLine(line, 380, 40, 430, 40);
            gfx.DrawLine(line, 430, 40, 430, 90);

            // Отрисовка рамки индекса слева вверху
            gfx.DrawLine(line, 118, 83, 218, 83);
            gfx.DrawLine(line, 118, 63, 118, 83);
            gfx.DrawLine(line, 218, 63, 218, 83);

            // Отрисовка рамки индекса справа снизу
            gfx.DrawLine(line, 230, 240, 230, 260);
            gfx.DrawLine(line, 330, 240, 330, 260);
            gfx.DrawLine(line, 330, 260, 230, 260);

            gfx.DrawString("От кого", font, XBrushes.Black,
                new XRect(10, 20, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(from, font2, XBrushes.Black,
                new XRect(70, 20, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("От куда", font, XBrushes.Black,
                new XRect(10, 40, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromAdress, font2, XBrushes.Black,
                new XRect(70, 40, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места отправления", fontIndex, XBrushes.Black,
                new XRect(120, 60, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromIndex, fontIndex2, XBrushes.Black,
                new XRect(150, 70, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Кому", font, XBrushes.Black,
                new XRect(200, 200, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(to, font2, XBrushes.Black,
                new XRect(240, 200, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Куда", font, XBrushes.Black,
                new XRect(200, 220, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toAdress, font2, XBrushes.Black,
                new XRect(240, 220, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места назначения", fontIndex, XBrushes.Black,
                new XRect(50, -75, page.Width, page.Height),
                XStringFormat.BottomCenter);
            gfx.DrawString(toIndex, fontIndex2, XBrushes.Black,
                new XRect(55, -65, page.Width, page.Height),
                XStringFormat.BottomCenter);

            string filename = "Test1.pdf";
            print.Save(filename);
            Process.Start(filename);

        }

        private void пустойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String from = textBox1.Text;
            String fromAdress = textBox2.Text;
            String fromIndex = textBox5.Text;
            String to = textBox3.Text;
            String toAdress = textBox4.Text;
            String toIndex = textBox6.Text;
            //Создание страницы (Конверт С5)
            PdfDocument print = new PdfDocument();
            PdfPage page = print.AddPage();
            page.Height = XUnit.FromMillimeter(162);
            page.Width = XUnit.FromMillimeter(229);
            XGraphics gfx = XGraphics.FromPdfPage(page);
            //Стили для документа
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFont font = new XFont("Times New Roman", 14, XFontStyle.Italic, options);
            XFont font2 = new XFont("Times New Roman", 14, XFontStyle.Underline, options);
            XFont fontIndex = new XFont("Times New Roman", 8, XFontStyle.Italic, options);
            XFont fontIndex2 = new XFont("Times New Roman", 10, XFontStyle.Bold, options);

            XPen line = new XPen(XColors.Black, 1);
            XPen lineForIndex = new XPen(XColors.Black, 5);
            XPen lineForIndex2 = new XPen(XColors.Black, 2);
            XPen DashLine = new XPen(XColors.Black, 1);
            DashLine.DashStyle = XDashStyle.Dash;

            // Отрисовка индекса слева внизу
            // Толстый пунктир
            gfx.DrawLine(lineForIndex, 140, 390, 160, 390);
            gfx.DrawLine(lineForIndex, 115, 390, 135, 390);
            gfx.DrawLine(lineForIndex, 90, 390, 110, 390);
            gfx.DrawLine(lineForIndex, 65, 390, 85, 390);
            gfx.DrawLine(lineForIndex, 40, 390, 60, 390);
            gfx.DrawLine(lineForIndex, 15, 390, 35, 390);
            gfx.DrawLine(lineForIndex, 165, 390, 185, 390);
                // Черточка слева
            gfx.DrawLine(lineForIndex2, 15, 400, 35, 400);
                // Крышечки сверху
            gfx.DrawLine(DashLine, 40, 400, 60, 400);
            gfx.DrawLine(DashLine, 65, 400, 85, 400);
            gfx.DrawLine(DashLine, 90, 400, 110, 400);
            gfx.DrawLine(DashLine, 115, 400, 135, 400);
            gfx.DrawLine(DashLine, 140, 400, 160, 400);
            gfx.DrawLine(DashLine, 165, 400, 185, 400);
                // Диагональные линии
            gfx.DrawLine(DashLine, 40, 420, 60, 400);
            gfx.DrawLine(DashLine, 40, 440, 60, 420);

            gfx.DrawLine(DashLine, 65, 420, 85, 400);
            gfx.DrawLine(DashLine, 65, 440, 85, 420);

            gfx.DrawLine(DashLine, 90, 420, 110, 400);
            gfx.DrawLine(DashLine, 90, 440, 110, 420);

            gfx.DrawLine(DashLine, 115, 420, 135, 400);
            gfx.DrawLine(DashLine, 115, 440, 135, 420);

            gfx.DrawLine(DashLine, 140, 420, 160, 400);
            gfx.DrawLine(DashLine, 140, 440, 160, 420);

            gfx.DrawLine(DashLine, 165, 420, 185, 400);
            gfx.DrawLine(DashLine, 165, 440, 185, 420);
                // Нижние крышечки
            gfx.DrawLine(DashLine, 40, 440, 60, 440);
            gfx.DrawLine(DashLine, 65, 440, 85, 440);
            gfx.DrawLine(DashLine, 90, 440, 110, 440);
            gfx.DrawLine(DashLine, 115, 440, 135, 440);
            gfx.DrawLine(DashLine, 140, 440, 160, 440);
            gfx.DrawLine(DashLine, 165, 440, 185, 440);
                // Средние крышечки
            gfx.DrawLine(DashLine, 40, 420, 60, 420);
            gfx.DrawLine(DashLine, 65, 420, 85, 420);
            gfx.DrawLine(DashLine, 90, 420, 110, 420);
            gfx.DrawLine(DashLine, 115, 420, 135, 420);
            gfx.DrawLine(DashLine, 140, 420, 160, 420);
            gfx.DrawLine(DashLine, 165, 420, 185, 420);
                //Палки слева
            gfx.DrawLine(DashLine, 40, 400, 40, 440);
            gfx.DrawLine(DashLine, 65, 400, 65, 440);
            gfx.DrawLine(DashLine, 90, 400, 90, 440);
            gfx.DrawLine(DashLine, 115, 400, 115, 440);
            gfx.DrawLine(DashLine, 140, 400, 140, 440);
            gfx.DrawLine(DashLine, 165, 400, 165, 440);
                //Палки справа
            gfx.DrawLine(DashLine, 60, 400, 60, 440);
            gfx.DrawLine(DashLine, 85, 400, 85, 440);
            gfx.DrawLine(DashLine, 110, 400, 110, 440);
            gfx.DrawLine(DashLine, 135, 400, 135, 440);
            gfx.DrawLine(DashLine, 160, 400, 160, 440);
            gfx.DrawLine(DashLine, 185, 400, 185, 440);

            // Отрисовка правого верхнего угла
            gfx.DrawLine(line, 500, 40, 600, 40);
            gfx.DrawLine(line, 600, 40, 600, 140);

            // Отрисовка рамки индекса слева вверху
            gfx.DrawLine(line, 140, 100, 250, 100);
            gfx.DrawLine(line, 140, 80, 140, 100);
            gfx.DrawLine(line, 250, 80, 250, 100);

            // Отрисовка рамки индекса справа снизу
            gfx.DrawLine(line, 350, 410, 460, 410);
            gfx.DrawLine(line, 350, 390, 350, 410);
            gfx.DrawLine(line, 460, 390, 460, 410);

            gfx.DrawString("От кого", font, XBrushes.Black,
                new XRect(40, 40, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(from, font2, XBrushes.Black,
                new XRect(100, 40, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("От куда", font, XBrushes.Black,
                new XRect(40, 60, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromAdress, font2, XBrushes.Black,
                new XRect(100, 60, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места отправления", fontIndex, XBrushes.Black,
                new XRect(150, 80, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromIndex, fontIndex2, XBrushes.Black,
                new XRect(185, 90, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Кому", font, XBrushes.Black,
                new XRect(335, 350, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(to, font2, XBrushes.Black,
                new XRect(370, 350, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Куда", font, XBrushes.Black,
                new XRect(335, 370, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toAdress, font2, XBrushes.Black,
                new XRect(370, 370, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места назначения", fontIndex, XBrushes.Black,
                new XRect(80, -60, page.Width, page.Height),
                XStringFormat.BottomCenter);
            gfx.DrawString(toIndex, fontIndex2, XBrushes.Black,
                new XRect(85, -50, page.Width, page.Height),
                XStringFormat.BottomCenter);

            string filename = "Test5.pdf";
            print.Save(filename);
            Process.Start(filename);
        }

        private void чистыйToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            String from = textBox1.Text;
            String fromAdress = textBox2.Text;
            String fromIndex = textBox5.Text;
            String to = textBox3.Text;
            String toAdress = textBox4.Text;
            String toIndex = textBox6.Text;
            //Создание страницы (Конверт С4)
            PdfDocument print = new PdfDocument();
            PdfPage page = print.AddPage();
            page.Height = XUnit.FromMillimeter(229);
            page.Width = XUnit.FromMillimeter(324);
            XGraphics gfx = XGraphics.FromPdfPage(page);
            //Стили для документа
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFont font = new XFont("Times New Roman", 14, XFontStyle.Italic, options);
            XFont font2 = new XFont("Times New Roman", 14, XFontStyle.Underline, options);
            XFont fontIndex = new XFont("Times New Roman", 8, XFontStyle.Italic, options);
            XFont fontIndex2 = new XFont("Times New Roman", 10, XFontStyle.Bold, options);

            XPen line = new XPen(XColors.Black, 1);
            XPen lineForIndex = new XPen(XColors.Black, 5);
            XPen lineForIndex2 = new XPen(XColors.Black, 2);
            XPen DashLine = new XPen(XColors.Black, 1);
            DashLine.DashStyle = XDashStyle.Dash;

            // Отрисовка индекса слева внизу
            // Толстый пунктир
            gfx.DrawLine(lineForIndex, 140, 560, 160, 560);
            gfx.DrawLine(lineForIndex, 115, 560, 135, 560);
            gfx.DrawLine(lineForIndex, 90, 560, 110, 560);
            gfx.DrawLine(lineForIndex, 65, 560, 85, 560);
            gfx.DrawLine(lineForIndex, 40, 560, 60, 560);
            gfx.DrawLine(lineForIndex, 15, 560, 35, 560);
            gfx.DrawLine(lineForIndex, 165, 560, 185, 560);
            // Черточка слева
            gfx.DrawLine(lineForIndex2, 15, 570, 35, 570);
            // Крышечки сверху
            gfx.DrawLine(DashLine, 40, 570, 60, 570);
            gfx.DrawLine(DashLine, 65, 570, 85, 570);
            gfx.DrawLine(DashLine, 90, 570, 110, 570);
            gfx.DrawLine(DashLine, 115, 570, 135, 570);
            gfx.DrawLine(DashLine, 140, 570, 160, 570);
            gfx.DrawLine(DashLine, 165, 570, 185, 570);
            // Диагональные линии
            gfx.DrawLine(DashLine, 40, 590, 60, 570);
            gfx.DrawLine(DashLine, 40, 610, 60, 590);

            gfx.DrawLine(DashLine, 65, 590, 85, 570);
            gfx.DrawLine(DashLine, 65, 610, 85, 590);

            gfx.DrawLine(DashLine, 90, 590, 110, 570);
            gfx.DrawLine(DashLine, 90, 610, 110, 590);

            gfx.DrawLine(DashLine, 115, 590, 135, 570);
            gfx.DrawLine(DashLine, 115, 610, 135, 590);

            gfx.DrawLine(DashLine, 140, 590, 160, 570);
            gfx.DrawLine(DashLine, 140, 610, 160, 590);

            gfx.DrawLine(DashLine, 165, 590, 185, 570);
            gfx.DrawLine(DashLine, 165, 610, 185, 590);
            // Нижние крышечки
            gfx.DrawLine(DashLine, 40, 610, 60, 610);
            gfx.DrawLine(DashLine, 65, 610, 85, 610);
            gfx.DrawLine(DashLine, 90, 610, 110, 610);
            gfx.DrawLine(DashLine, 115, 610, 135, 610);
            gfx.DrawLine(DashLine, 140, 610, 160, 610);
            gfx.DrawLine(DashLine, 165, 610, 185, 610);
            // Средние крышечки
            gfx.DrawLine(DashLine, 40, 590, 60, 590);
            gfx.DrawLine(DashLine, 65, 590, 85, 590);
            gfx.DrawLine(DashLine, 90, 590, 110, 590);
            gfx.DrawLine(DashLine, 115, 590, 135, 590);
            gfx.DrawLine(DashLine, 140, 590, 160, 590);
            gfx.DrawLine(DashLine, 165, 590, 185, 590);
            //Палки слева
            gfx.DrawLine(DashLine, 40, 570, 40, 610);
            gfx.DrawLine(DashLine, 65, 570, 65, 610);
            gfx.DrawLine(DashLine, 90, 570, 90, 610);
            gfx.DrawLine(DashLine, 115, 570, 115, 610);
            gfx.DrawLine(DashLine, 140, 570, 140, 610);
            gfx.DrawLine(DashLine, 165, 570, 165, 610);
            //Палки справа
            gfx.DrawLine(DashLine, 60, 570, 60, 610);
            gfx.DrawLine(DashLine, 85, 570, 85, 610);
            gfx.DrawLine(DashLine, 110, 570, 110, 610);
            gfx.DrawLine(DashLine, 135, 570, 135, 610);
            gfx.DrawLine(DashLine, 160, 570, 160, 610);
            gfx.DrawLine(DashLine, 185, 570, 185, 610);

            // Отрисовка правого верхнего угла
            gfx.DrawLine(line, 750, 80, 850, 80);
            gfx.DrawLine(line, 850, 80, 850, 180);

            // Отрисовка рамки индекса слева вверху
            gfx.DrawLine(line, 140, 165, 250, 165);
            gfx.DrawLine(line, 140, 145, 140, 165);
            gfx.DrawLine(line, 250, 145, 250, 165);

            // Отрисовка рамки индекса справа снизу
            gfx.DrawLine(line, 585, 560, 685, 560);
            gfx.DrawLine(line, 585, 540, 585, 560);
            gfx.DrawLine(line, 685, 540, 685, 560);

            gfx.DrawString("От кого", font, XBrushes.Black,
                new XRect(60, 100, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(from, font2, XBrushes.Black,
                new XRect(120, 100, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("От куда", font, XBrushes.Black,
                new XRect(60, 120, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromAdress, font2, XBrushes.Black,
                new XRect(120, 120, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места отправления", fontIndex, XBrushes.Black,
                new XRect(150, 140, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromIndex, fontIndex2, XBrushes.Black,
                new XRect(180, 150, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Кому", font, XBrushes.Black,
                new XRect(550, 500, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(to, font2, XBrushes.Black,
                new XRect(590, 500, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Куда", font, XBrushes.Black,
                new XRect(550, 520, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toAdress, font2, XBrushes.Black,
                new XRect(590, 520, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места назначения", fontIndex, XBrushes.Black,
                new XRect(590, 540, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toIndex, fontIndex2, XBrushes.Black,
                new XRect(620, 550, page.Width, page.Height),
                XStringFormat.TopLeft);

            string filename = "Test23.pdf";
            print.Save(filename);
            Process.Start(filename);
        }

        private void чистыйToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            String from = textBox1.Text;
            String fromAdress = textBox2.Text;
            String fromIndex = textBox5.Text;
            String to = textBox3.Text;
            String toAdress = textBox4.Text;
            String toIndex = textBox6.Text;
            //Создание страницы (Конверт DL "E65")
            PdfDocument print = new PdfDocument();
            PdfPage page = print.AddPage();
            page.Height = XUnit.FromMillimeter(110);
            page.Width = XUnit.FromMillimeter(229);
            XGraphics gfx = XGraphics.FromPdfPage(page);
            //Стили для документа
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFont font = new XFont("Times New Roman", 14, XFontStyle.Italic, options);
            XFont font2 = new XFont("Times New Roman", 14, XFontStyle.Underline, options);
            XFont fontIndex = new XFont("Times New Roman", 8, XFontStyle.Italic, options);
            XFont fontIndex2 = new XFont("Times New Roman", 10, XFontStyle.Bold, options);

            XPen line = new XPen(XColors.Black, 1);
            XPen lineForIndex = new XPen(XColors.Black, 5);
            XPen lineForIndex2 = new XPen(XColors.Black, 2);
            XPen DashLine = new XPen(XColors.Black, 1);
            DashLine.DashStyle = XDashStyle.Dash;

            // Отрисовка индекса слева внизу
            // Толстый пунктир
            gfx.DrawLine(lineForIndex, 140, 220, 160, 220);
            gfx.DrawLine(lineForIndex, 115, 220, 135, 220);
            gfx.DrawLine(lineForIndex, 90, 220, 110, 220);
            gfx.DrawLine(lineForIndex, 65, 220, 85, 220);
            gfx.DrawLine(lineForIndex, 40, 220, 60, 220);
            gfx.DrawLine(lineForIndex, 15, 220, 35, 220);
            gfx.DrawLine(lineForIndex, 165, 220, 185, 220);
            // Черточка слева
            gfx.DrawLine(lineForIndex2, 15, 230, 35, 230);
            // Крышечки сверху
            gfx.DrawLine(DashLine, 40, 230, 60, 230);
            gfx.DrawLine(DashLine, 65, 230, 85, 230);
            gfx.DrawLine(DashLine, 90, 230, 110, 230);
            gfx.DrawLine(DashLine, 115, 230, 135, 230);
            gfx.DrawLine(DashLine, 140, 230, 160, 230);
            gfx.DrawLine(DashLine, 165, 230, 185, 230);
            // Диагональные линии
            gfx.DrawLine(DashLine, 40, 250, 60, 230);
            gfx.DrawLine(DashLine, 40, 270, 60, 250);

            gfx.DrawLine(DashLine, 65, 250, 85, 230);
            gfx.DrawLine(DashLine, 65, 270, 85, 250);

            gfx.DrawLine(DashLine, 90, 250, 110, 230);
            gfx.DrawLine(DashLine, 90, 270, 110, 250);

            gfx.DrawLine(DashLine, 115, 250, 135, 230);
            gfx.DrawLine(DashLine, 115, 270, 135, 250);

            gfx.DrawLine(DashLine, 140, 250, 160, 230);
            gfx.DrawLine(DashLine, 140, 270, 160, 250);

            gfx.DrawLine(DashLine, 165, 250, 185, 230);
            gfx.DrawLine(DashLine, 165, 270, 185, 250);
            // Нижние крышечки
            gfx.DrawLine(DashLine, 40, 270, 60, 270);
            gfx.DrawLine(DashLine, 65, 270, 85, 270);
            gfx.DrawLine(DashLine, 90, 270, 110, 270);
            gfx.DrawLine(DashLine, 115, 270, 135, 270);
            gfx.DrawLine(DashLine, 140, 270, 160, 270);
            gfx.DrawLine(DashLine, 165, 270, 185, 270);
            // Средние крышечки
            gfx.DrawLine(DashLine, 40, 250, 60, 250);
            gfx.DrawLine(DashLine, 65, 250, 85, 250);
            gfx.DrawLine(DashLine, 90, 250, 110, 250);
            gfx.DrawLine(DashLine, 115, 250, 135, 250);
            gfx.DrawLine(DashLine, 140, 250, 160, 250);
            gfx.DrawLine(DashLine, 165, 250, 185, 250);
            //Палки слева
            gfx.DrawLine(DashLine, 40, 230, 40, 270);
            gfx.DrawLine(DashLine, 65, 230, 65, 270);
            gfx.DrawLine(DashLine, 90, 230, 90, 270);
            gfx.DrawLine(DashLine, 115, 230, 115, 270);
            gfx.DrawLine(DashLine, 140, 230, 140, 270);
            gfx.DrawLine(DashLine, 165, 230, 165, 270);
            //Палки справа
            gfx.DrawLine(DashLine, 60, 230, 60, 270);
            gfx.DrawLine(DashLine, 85, 230, 85, 270);
            gfx.DrawLine(DashLine, 110, 230, 110, 270);
            gfx.DrawLine(DashLine, 135, 230, 135, 270);
            gfx.DrawLine(DashLine, 160, 230, 160, 270);
            gfx.DrawLine(DashLine, 185, 230, 185, 270);

            // Отрисовка правого верхнего угла
            gfx.DrawLine(line, 560, 30, 610, 30);
            gfx.DrawLine(line, 610, 30, 610, 80);

            // Отрисовка рамки индекса слева вверху
            gfx.DrawLine(line, 137, 80, 247, 80);
            gfx.DrawLine(line, 137, 60, 137, 80);
            gfx.DrawLine(line, 247, 60, 247, 80);

            // Отрисовка рамки индекса справа снизу
            gfx.DrawLine(line, 443, 210, 443, 230);
            gfx.DrawLine(line, 543, 210, 543, 230);
            gfx.DrawLine(line, 443, 230, 543, 230);

            gfx.DrawString("От кого", font, XBrushes.Black,
                new XRect(20, 20, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(from, font2, XBrushes.Black,
                new XRect(80, 20, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("От куда", font, XBrushes.Black,
                new XRect(20, 40, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromAdress, font2, XBrushes.Black,
                new XRect(80, 40, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места отправления", fontIndex, XBrushes.Black,
                new XRect(145, 60, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(fromIndex, fontIndex2, XBrushes.Black,
                new XRect(175, 70, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Кому", font, XBrushes.Black,
                new XRect(350, 170, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(to, font2, XBrushes.Black,
                new XRect(390, 170, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Куда", font, XBrushes.Black,
                new XRect(350, 190, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toAdress, font2, XBrushes.Black,
                new XRect(390, 190, page.Width, page.Height),
                XStringFormat.TopLeft);

            gfx.DrawString("Индекс места назначения", fontIndex, XBrushes.Black,
                new XRect(450, 210, page.Width, page.Height),
                XStringFormat.TopLeft);
            gfx.DrawString(toIndex, fontIndex2, XBrushes.Black,
                new XRect(480, 220, page.Width, page.Height),
                XStringFormat.TopLeft);

            string filename = "Test3.pdf";
            print.Save(filename);
            Process.Start(filename);
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = "InfoCmon.chm";
            Process.Start(filename);
        }

        private static string WebRequest(string from)
        {
            const string WEBSERVICE_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address";
            try
            {
                var webRequest = System.Net.WebRequest.Create(WEBSERVICE_URL);
                if (webRequest != null)
                {
                    webRequest.Method = "POST";
                    webRequest.Timeout = 20000;
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add("Authorization", "Token de9cb0ec7c3251c64d06a7a823bc47fdba595a11");
                    //String str = String.Format("{ \"query\": \"{0} \", \"count\": 1}", from);
                    //String str = "\"{ \"query\": \"Новосибирск, Кирова \", \"count\": 1}\"";
                    string test = from;
                    string requestString = "{ \"query\": \"" + from + "\" , \"count\": 1}";

                    string json = requestString;
                    using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
                    {
                        streamWriter.Write(json);
                        streamWriter.Flush();
                    }

                    var httpResponse = (HttpWebResponse)webRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var responseText = streamReader.ReadToEnd();
                        var vr1 = 0;
                        dynamic stuff = Newtonsoft.Json.JsonConvert.DeserializeObject(responseText);

                        foreach (var item in stuff.suggestions)
                        {
                            Console.WriteLine(item.data.postal_code);
                            vr1 = item.data.postal_code;
                        }
                        string fromItem = Convert.ToString(vr1);
                        return fromItem;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                return null;
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            string adress = textBox2.Text;
            string indexFromAdress = WebRequest(adress);
            textBox5.Text = indexFromAdress;
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            string adress = textBox4.Text;
            string indexFromAdress = WebRequest(adress);
            textBox6.Text = indexFromAdress;
        }
    }
}
