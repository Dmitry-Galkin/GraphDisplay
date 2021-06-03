using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using ZedGraph;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;



namespace GraphDisplay
{
    public partial class FormGraph : Form
    {
        // массив данных для построения
        double[][] plotData;
        double[] X, Y;
        // массив строк для отображаемых данных
        string[] stringData;
        // панель для рисования
        GraphPane pane;
        // текущий график
        LineItem plot;
        // цвет линии
        Color color;
        // диреткория сохраняемого excel-файла с переменными
        string path_file_xlsx = null;
        // массив цветов линий
        string[] strLineColor = { "Красный", "Синий", "Зеленый", "Черный", "Золотой", "Желтый"};
        // массив стилей линий
        string[] strLineStyle = { "Сплошная", "Штрихпунктирная", "Круг", "Квадрат", "Ромб", "Треугольник" };
        // массивы ComboBox
        ComboBox[] cBoxData = new ComboBox[7];
        ComboBox[] cBoxLineStyle = new ComboBox[6];
        ComboBox[] cBoxLineColor = new ComboBox[6];
        ComboBox[] cBoxLineWidth = new ComboBox[6];
        // массив TextBox
        TextBox[] txtBoxCoefScaleY = new TextBox[6];
        // массив CheckBox
        CheckBox[] chBoxDataY = new CheckBox[6];
        // коэф-т масштабирования по оси Y
        double coefScaleY;
        // флаги установки границ по осям
        bool limX, limY;
        // минимальное и максимальное значения загруженных массивов
        double minValue, maxValue;
        // промежуточные переменные
        int L, n, m, n_max, Nplot, kRarity;
        string str;




        public FormGraph()
        {
            InitializeComponent();

            // подпишемся на событие, которое будет возникать перед тем, 
            // как будет показано контекстное меню
            zedGraphPane.ContextMenuBuilder +=
                new ZedGraphControl.ContextMenuBuilderEventHandler(zedGraph_ContextMenuBuilder);

            // начальное состояние формы
            initialStateForm();


            /*int N = 7001;
            double[][] data = new double[4][];
            double[] T = new double[N];
            double[] x = new double[N];
            double[] y = new double[N];
            double[] z = new double[N];
            
            for (int i = 0; i < N; i++)
            {
                T[i] = 0.1 * i;
                x[i] = 5 * Math.Sin(0.05 * T[i]);
                y[i] = 5 * Math.Cos(0.1 * T[i]);
                if (i < (N - 10)) z[i] = 0.005 * i;
            }
            data[0] = x;
            data[1] = y;
            data[2] = z;
            data[3] = T;
            string[] strPlot = { "x", "y", "z", "T" };

            Initialization(data, strPlot);*/
        }


        // ОБРАБОТЧИК СОБЫТИЯ, КОТОРЫЙ ВЫЗЫВАЕТСЯ ПЕРЕД ПОКАЗОМ КОНТЕКСТНОГО МЕНЮ
        void zedGraph_ContextMenuBuilder(ZedGraphControl sender,
            ContextMenuStrip menuStrip,
            Point mousePt,
            ZedGraphControl.ContextMenuObjectState objState)
        {
            // удаляем пункты меню...
            menuStrip.Items.RemoveAt(2); // ..."параметры страницы"
            menuStrip.Items.RemoveAt(2); // ..."печать"
            menuStrip.Items.RemoveAt(3); // ..."отменить последнее масштабирование"
        }


        // НАЧАЛЬНОЕ СОСТОЯНИЕ ФОРМЫ
        private void initialStateForm()
        {
            // начальные размеры формы (90% от текущего экрана)
            Rectangle resolutionRect = Screen.PrimaryScreen.Bounds;
            this.Height = resolutionRect.Height * 90 / 100;
            this.Width = resolutionRect.Width * 90 / 100;

            // кнопка отрисовки в нач. момент недоступна
            this.buttonDraw.Enabled = false;

            // получим панель для рисования
            pane = zedGraphPane.GraphPane;

            // панель для рисования делаем на переднем фронт
            this.zedGraphPane.BringToFront();

            // кнопка "сужения" в нач. момент невидимая
            this.buttonNarrowZGPane.Visible = false;

            // массив из ComboBox для задания цвета линий
            cBoxLineColor[0] = cBoxLineColor1;
            cBoxLineColor[1] = cBoxLineColor2;
            cBoxLineColor[2] = cBoxLineColor3;
            cBoxLineColor[3] = cBoxLineColor4;
            cBoxLineColor[4] = cBoxLineColor5;
            cBoxLineColor[5] = cBoxLineColor6;

            // массив из ComboBox для задания стиля линий
            cBoxLineStyle[0] = cBoxLineStyle1;
            cBoxLineStyle[1] = cBoxLineStyle2;
            cBoxLineStyle[2] = cBoxLineStyle3;
            cBoxLineStyle[3] = cBoxLineStyle4;
            cBoxLineStyle[4] = cBoxLineStyle5;
            cBoxLineStyle[5] = cBoxLineStyle6;

            // массив из ComboBox для задания толщины линий
            cBoxLineWidth[0] = cBoxLineWidth1;
            cBoxLineWidth[1] = cBoxLineWidth2;
            cBoxLineWidth[2] = cBoxLineWidth3;
            cBoxLineWidth[3] = cBoxLineWidth4;
            cBoxLineWidth[4] = cBoxLineWidth5;
            cBoxLineWidth[5] = cBoxLineWidth6;

            // массив из TextBox для задания коэф-та масштабирования по оси Y
            txtBoxCoefScaleY[0] = textBoxCoefScaleY1;
            txtBoxCoefScaleY[1] = textBoxCoefScaleY2;
            txtBoxCoefScaleY[2] = textBoxCoefScaleY3;
            txtBoxCoefScaleY[3] = textBoxCoefScaleY4;
            txtBoxCoefScaleY[4] = textBoxCoefScaleY5;
            txtBoxCoefScaleY[5] = textBoxCoefScaleY6;

            // массив из ComboBox для отображаемых данных
            cBoxData[0] = cBoxDataY1;
            cBoxData[1] = cBoxDataY2;
            cBoxData[2] = cBoxDataY3;
            cBoxData[3] = cBoxDataY4;
            cBoxData[4] = cBoxDataY5;
            cBoxData[5] = cBoxDataY6;
            cBoxData[6] = cBoxDataX;

            // массив из CheckBox для отображения данных
            chBoxDataY[0] = checkBoxDataY1;
            chBoxDataY[1] = checkBoxDataY2;
            chBoxDataY[2] = checkBoxDataY3;
            chBoxDataY[3] = checkBoxDataY4;
            chBoxDataY[4] = checkBoxDataY5;
            chBoxDataY[5] = checkBoxDataY6;

            // начальные значения CheckBox для сетки
            checkBoxMajorGridX.Checked = true;
            checkBoxMajorGridY.Checked = true;
            if (checkBoxMajorGridX.Checked) pane.XAxis.MajorGrid.IsVisible = true;
            else pane.XAxis.MajorGrid.IsVisible = false;
            if (checkBoxMajorGridY.Checked) pane.YAxis.MajorGrid.IsVisible = true;
            else pane.YAxis.MajorGrid.IsVisible = false;

            // начальные значения CheckBox для центральных осей
            checkBoxAxisCentralX.Checked = false;
            checkBoxAxisCentralY.Checked = false;
            if (checkBoxAxisCentralX.Checked) pane.YAxis.MajorGrid.IsZeroLine = true;
            else pane.YAxis.MajorGrid.IsZeroLine = false;
            if (checkBoxAxisCentralY.Checked) pane.XAxis.MajorGrid.IsZeroLine = true;
            else pane.XAxis.MajorGrid.IsZeroLine = false;

            // начальные значения CheckBox для отображения графики
            checkBoxDataY1.Checked = true;
            checkBoxDataY2.Checked = true;
            checkBoxDataY3.Checked = true;
            checkBoxDataY4.Checked = false;
            checkBoxDataY5.Checked = false;
            checkBoxDataY6.Checked = false;

            // нач. значение CheckBox для легенды
            checkBoxLegend.Checked = true;
            // положение легенды
            pane.Legend.Position = LegendPos.InsideTopRight;
            // шрифт легенды
            //pane.Legend.FontSpec.Size = 14;


            // надпись над графиком
            pane.Title.Text = "";
            // текст надписи по оси X
            pane.XAxis.Title.Text = "";
            // текст надписи по оси Y
            pane.YAxis.Title.Text = "";
            // размер шрифта по осям для сетки
            pane.XAxis.Scale.FontSpec.Size = 14;
            pane.XAxis.Title.FontSpec.Size = 14;
            pane.YAxis.Scale.FontSpec.Size = 14;
            pane.YAxis.Title.FontSpec.Size = 14;
            // запрет на изменение шрифтов при масштабировании
            pane.IsFontsScaled = false;


            // рамку графики делаем светлой, чтобы при копировании в отчет не мешалась
            pane.Border.Color = Color.White;


            // заполнение ComboBox...
            for (int i = 0; i < 6; i++)
            {
                //...для задания стиля линий
                L = strLineStyle.Length;
                for (int j = 0; j < L; j++) cBoxLineStyle[i].Items.Add(strLineStyle[j]);
                cBoxLineStyle[i].SelectedIndex = 0;

                //...для задания цвета линий
                L = strLineColor.Length;
                for (int j = 0; j < L; j++) cBoxLineColor[i].Items.Add(strLineColor[j]);
                try { cBoxLineColor[i].SelectedIndex = i; }
                catch { };

                //...для задания толщины линий
                for (int j = 1; j <= 12; j++) cBoxLineWidth[i].Items.Add(j.ToString());
                cBoxLineWidth[i].SelectedIndex = 2;

                //...для задания коэф-та масштабирования по оси Y
                txtBoxCoefScaleY[i].Text = "1,0";
            }


            // TextBox с коэффициентом разреженности вывода графики
            textBoxRarity.Text = "1";
            // TextBox с размером шрифта для легенды
            textBox_sizeTextLegend.Text = "14";
            // TextBox с размером шрифта для подписей
            textBox_sizeTextAxis.Text = "14";


            // Настраиваем DataGridView
            dataGridViewAnalysis.RowHeadersVisible = false;
            dataGridViewAnalysis.Columns.Add("variable_name", "Name");
            dataGridViewAnalysis.Columns.Add("variable_min", "Min");
            dataGridViewAnalysis.Columns.Add("variable_max", "Max");
            dataGridViewAnalysis.Columns.Add("variable_size", "Size");
            dataGridViewAnalysis.Columns[0].Width = (int)(dataGridViewAnalysis.Width / 4.01);
            dataGridViewAnalysis.Columns[1].Width = (int)(dataGridViewAnalysis.Width / 4.01);
            dataGridViewAnalysis.Columns[2].Width = (int)(dataGridViewAnalysis.Width / 4.01);
            dataGridViewAnalysis.Columns[3].Width = (int)(dataGridViewAnalysis.Width / 4.01);
            dataGridViewAnalysis.AllowUserToAddRows = false;
            dataGridViewAnalysis.AllowUserToResizeRows = false;
            //dataGridViewAnalysis.AllowUserToResizeColumns = false;
        }


        // ИНИЦИАЛИЗАЦИЯ
        public void Initialization(double[][] PlotData, string[] strPlotData)
        {
            this.plotData = PlotData;
            this.stringData = strPlotData;

            // заполнение всех ComboBox данными для отображения
            L = strPlotData.Length;
            for (int i = 0; i < 7; i++)
            {
                for (int j = 0; j < L; j++)
                    cBoxData[i].Items.Add(strPlotData[j]);

                if (i == 6) cBoxData[i].SelectedIndex = L - 1;
                else
                {
                    if (i != (L - 1))
                    {
                        try { cBoxData[i].SelectedIndex = i; }
                        catch { }
                    }
                }
            }

            // заполняем DataGridView
            for (int i = 0; i < L; i++)
            {
                // название переменной
                dataGridViewAnalysis.Rows.Add(strPlotData[i]);
                // минимальное значение
                dataGridViewAnalysis[1, i].Value = minValueArray(plotData[i]).ToString();
                // максимальное значение
                dataGridViewAnalysis[2, i].Value = maxValueArray(plotData[i]).ToString();
                // размер массива
                dataGridViewAnalysis[3, i].Value = (plotData[i].Length).ToString();
            }

            // делаем доступной кнопку отрисовки графики
            buttonDraw.Enabled = true;
        }


        // ЗАДАНИЕ ЦВЕТА ЛИНИИ
        private void lineColor(int numCBoxLineColor)
        {
            switch (numCBoxLineColor)
            {
                // красный
                case 0:
                    color = Color.Red;
                    break;
                // синий
                case 1:
                    color = Color.Blue;
                    break;
                // зеленый
                case 2:
                    color = Color.Green;
                    break;
                // черный
                case 3:
                    color = Color.Black;
                    break;
                // золотой
                case 4:
                    color = Color.Gold;
                    break;
                // желтый
                case 5:
                    color = Color.Yellow;
                    break;
                default:
                    break;
            }
        }


        // ЗАДАНИЕ СТИЛЯ ЛИНИИ
        private void lineStyleWidth(int numCBoxLineStyle, int lineWidth)
        {
            switch (numCBoxLineStyle)
            {
                // сплошная линия
                case 0:
                    plot.Line.IsVisible = true;
                    plot.Line.Width = lineWidth;
                    plot.Symbol.IsVisible = false;
                    break;
                // штрихпунктирная линия
                case 1:
                    plot.Line.IsVisible = true;
                    plot.Line.Style = DashStyle.Dash;
                    plot.Line.Width = lineWidth;
                    plot.Line.Style = DashStyle.Custom;
                    plot.Line.DashOn = 7.0f;
                    plot.Line.DashOff = 3.0f;
                    plot.Line.IsSmooth = true;
                    plot.Symbol.IsVisible = false;
                    break;
                // маркер - круг
                case 2:
                    plot.Line.IsVisible = false;
                    plot.Symbol.IsVisible = true;
                    plot.Symbol.Type = SymbolType.Circle;
                    plot.Symbol.Fill.Color = color;
                    plot.Symbol.Fill.Type = FillType.Solid;
                    plot.Symbol.Size = lineWidth;
                    break;
                // маркер - квадрат
                case 3:
                    plot.Line.IsVisible = false;
                    plot.Symbol.IsVisible = true;
                    plot.Symbol.Type = SymbolType.Square;
                    plot.Symbol.Fill.Color = color;
                    plot.Symbol.Fill.Type = FillType.Solid;
                    plot.Symbol.Size = lineWidth;
                    break;
                // маркер - ромб
                case 4:
                    plot.Line.IsVisible = false;
                    plot.Symbol.IsVisible = true;
                    plot.Symbol.Type = SymbolType.Diamond;
                    plot.Symbol.Fill.Color = color;
                    plot.Symbol.Fill.Type = FillType.Solid;
                    plot.Symbol.Size = lineWidth;
                    break;
                // маркер - треугольник
                case 5:
                    plot.Line.IsVisible = false;
                    plot.Symbol.IsVisible = true;
                    plot.Symbol.Type = SymbolType.Triangle;
                    plot.Symbol.Fill.Color = color;
                    plot.Symbol.Fill.Type = FillType.Solid;
                    plot.Symbol.Size = lineWidth;
                    break;
                default:
                    break;
            }
        }


        // ОСНОВНАЯ СЕТКА ПО ОСИ X
        private void checkBoxMajorGridX_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMajorGridX.Checked)
            {
                pane.XAxis.MajorGrid.IsVisible = true;
                pane.XAxis.MajorGrid.DashOn = 5;
                pane.XAxis.MajorGrid.DashOff = 5;
            }
            else pane.XAxis.MajorGrid.IsVisible = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // ОСНОВНАЯ СЕТКА ПО ОСИ Y
        private void checkBoxMajorGridY_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMajorGridY.Checked)
            {
                pane.YAxis.MajorGrid.IsVisible = true;
                pane.YAxis.MajorGrid.DashOn = 5;
                pane.YAxis.MajorGrid.DashOff = 5;
            }
            else pane.YAxis.MajorGrid.IsVisible = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // ВСПОМОГАТЕЛЬНАЯ СЕТКА ПО ОСИ X
        private void checkBoxMinorGridX_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMinorGridX.Checked)
            {
                pane.XAxis.MinorGrid.IsVisible = true;
                pane.XAxis.MinorGrid.DashOn = 2;
                pane.XAxis.MinorGrid.DashOff = 2;
            }
            else pane.XAxis.MinorGrid.IsVisible = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // ВСПОМОГАТЕЛЬНАЯ СЕТКА ПО ОСИ Y
        private void checkBoxMinorGridY_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMinorGridY.Checked)
            {
                pane.YAxis.MinorGrid.IsVisible = true;
                pane.YAxis.MinorGrid.DashOn = 2;
                pane.YAxis.MinorGrid.DashOff = 2;
            }
            else pane.YAxis.MinorGrid.IsVisible = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // ЦЕНТРАЛЬНАЯ ОСЬ X
        private void checkBoxAxisCentralX_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAxisCentralX.Checked) pane.YAxis.MajorGrid.IsZeroLine = true;
            else pane.YAxis.MajorGrid.IsZeroLine = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // ЦЕНТРАЛЬНАЯ ОСЬ Y
        private void checkBoxAxisCentralY_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAxisCentralY.Checked) pane.XAxis.MajorGrid.IsZeroLine = true;
            else pane.XAxis.MajorGrid.IsZeroLine = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // КНОПКА "ВЫХОД" НА ФОРМЕ
        private void buttonExit_Click(object sender, EventArgs e)
        {
            Close();
        }


        // ПОСТРОЕНИЕ ГРАФИКИ
        private void buttonDraw_Click(object sender, EventArgs e)
        {
            pane.GraphObjList.Clear();
            // очистим список кривых
            pane.CurveList.Clear();
            // коэф-т разреженности вывода графики
            try { kRarity = Convert.ToInt32(textBoxRarity.Text); }
            catch { kRarity = 1; }


            for (int i = 0; i < 6; i++)
            {
                if (chBoxDataY[i].Checked)
                {
                    // коэф-т масштабирования по оси Y
                    try { coefScaleY = Convert.ToDouble(txtBoxCoefScaleY[i].Text); }
                    catch { coefScaleY = 1.0; }

                    // устанавливаем цвет текущего графика
                    lineColor(cBoxLineColor[i].SelectedIndex);

                    // отрисовка графика
                    if (kRarity != 1)
                    {
                        L = this.plotData[cBoxDataX.SelectedIndex].Length;
                        if (L != this.plotData[cBoxData[i].SelectedIndex].Length)
                        {
                            MessageBox.Show("Массивы должны иметь одинаковую размерность");
                            break;
                        }
                        if ((L % kRarity == 0) || (L % kRarity == 1)) Nplot = L / kRarity + 1;
                        else Nplot = L / kRarity + 2;
                        X = new double[Nplot];
                        Y = new double[Nplot];
                        n = 0;
                        // прореживаем точки
                        for (int k = 0; k < L; k++)
                        {
                            if (((k % kRarity) == 0) || (k == (L - 1)))
                            {
                                X[n] = this.plotData[cBoxDataX.SelectedIndex][k];
                                Y[n] = coefScaleY * this.plotData[cBoxData[i].SelectedIndex][k];
                                n++;
                            }
                        }
                    }
                    else
                    {
                        L = this.plotData[cBoxDataX.SelectedIndex].Length;
                        if (L != this.plotData[cBoxData[i].SelectedIndex].Length)
                        {
                            MessageBox.Show("Массивы должны иметь одинаковую размерность");
                            break;
                        }
                        X = new double[L];
                        Y = new double[L];
                        for (int k = 0; k < L; k++)
                        {
                            X[k] = this.plotData[cBoxDataX.SelectedIndex][k];
                            Y[k] = coefScaleY * this.plotData[cBoxData[i].SelectedIndex][k];
                        }
                    }

                    // строим график из прореженных точек
                    plot = pane.AddCurve(this.stringData[cBoxData[i].SelectedIndex],
                                            X,
                                            Y,
                                            color);



                    // устанавливаем тип и толщину линии
                    lineStyleWidth(cBoxLineStyle[i].SelectedIndex, cBoxLineWidth[i].SelectedIndex + 1);              
                }
            }


            // размер шрифта легенды
            try
            {
                pane.Legend.FontSpec.Size = Convert.ToInt16(textBox_sizeTextLegend.Text);
            }
            catch
            {
                pane.Legend.FontSpec.Size = 14;
            }


            // размер шрифта подписей по осям
            try
            {
                pane.XAxis.Scale.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
                pane.XAxis.Title.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
                pane.YAxis.Scale.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
                pane.YAxis.Title.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
            }
            catch
            {
                pane.XAxis.Scale.FontSpec.Size = 14;
                pane.XAxis.Title.FontSpec.Size = 14;
                pane.YAxis.Scale.FontSpec.Size = 14;
                pane.YAxis.Title.FontSpec.Size = 14;
            }


            // устанавливаем границы по оси X
            try
            {
                pane.XAxis.Scale.Min = Convert.ToDouble(textBoxXminLim.Text);
                pane.XAxis.Scale.Max = Convert.ToDouble(textBoxXmaxLim.Text);
                limX = true;

                if (checkBoxManualXTickLeft.Checked)
                {
                    TextObj text = new TextObj(textBoxXminLim.Text, 
                        pane.XAxis.Scale.Min, pane.YAxis.Scale.Min - Convert.ToDouble(textBox_TicksShift.Text));
                    text.Location.AlignH = AlignH.Center;
                    text.Location.AlignV = AlignV.Top;
                    text.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
                    text.FontSpec.Border.IsVisible = false;
                    text.FontSpec.Fill.IsVisible = false;
                    pane.GraphObjList.Add(text);
                }

                if (checkBoxManualYTickRight.Checked)
                {
                    TextObj text = new TextObj(textBoxXmaxLim.Text,
                        pane.XAxis.Scale.Max, pane.YAxis.Scale.Min - Convert.ToDouble(textBox_TicksShift.Text));
                    text.Location.AlignH = AlignH.Center;
                    text.Location.AlignV = AlignV.Top;
                    text.FontSpec.Size = Convert.ToInt16(textBox_sizeTextAxis.Text);
                    text.FontSpec.Border.IsVisible = false;
                    text.FontSpec.Fill.IsVisible = false;
                    pane.GraphObjList.Add(text);
                }
            }
            catch { limX = false; }
            // устанавливаем границы по оси Y
            try
            {
                pane.YAxis.Scale.Min = Convert.ToDouble(textBoxYminLim.Text);
                pane.YAxis.Scale.Max = Convert.ToDouble(textBoxYmaxLim.Text);
                limY = true;
            }
            catch { limY = false; }


            // автоматический масштаб по осям
            if ((limX) && (!limY))
            {
                pane.YAxis.Scale.MinAuto = true;
                pane.YAxis.Scale.MaxAuto = true;
            }
            else if ((limY) && (!limX))
            {
                pane.XAxis.Scale.MinAuto = true;
                pane.XAxis.Scale.MaxAuto = true;
            }
            else if ((!limY) && (!limX))
            {
                pane.XAxis.Scale.MinAuto = true;
                pane.XAxis.Scale.MaxAuto = true;
                pane.YAxis.Scale.MinAuto = true;
                pane.YAxis.Scale.MaxAuto = true;
            }
            
            // учет только видимого интервала графика
            pane.IsBoundedRanges = true;

            // показатель степени коэф-та умножения данных по осям (10^...)
            pane.XAxis.Scale.Mag = Convert.ToInt32(textBoxCoefMagX.Text);
            pane.YAxis.Scale.Mag = Convert.ToInt32(textBoxCoefMagY.Text);

            // обновляем данные об осях
            zedGraphPane.AxisChange();
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // РАСТЯГИВАНИЕ ПАНЕЛИ ДЛЯ РИСОВАНИЯ
        private void buttonWideZGPane_Click(object sender, EventArgs e)
        {
            // растягиваем панель рисования на всю форму
            zedGraphPane.Width = this.Width - 4 * this.zedGraphPane.Location.X -this.buttonWideZGPane.Width;
            // оставшиеся видимыми элементы делаем невидимыми
            for (int i = 0; i < 6; i++)
            {
                cBoxLineWidth[i].Visible = false;
                txtBoxCoefScaleY[i].Visible = false;
            }
            labelLineWidth.Visible = false;
            labelCoefScaleY.Visible = false;
            //textBox_sizeTextAxis.Visible = false;
            groupBoxAxisCentral.Visible = false;
            groupBox_outParameter.Visible = false;
            dataGridViewAnalysis.Visible = false;

            buttonNarrowZGPane.Visible = true;
        }


        // СУЖЕНИЕ ПАНЕЛИ ДЛЯ РИСОВАНИЯ
        private void buttonNarrowZGPane_Click(object sender, EventArgs e)
        {
            // сужаем панель рисования до начального состояния
            zedGraphPane.Width = this.checkBoxDataY1.Location.X - 3 * this.zedGraphPane.Location.X;
            // невидимые элементы делаем видимыми
            for (int i = 0; i < 6; i++)
            {
                cBoxLineWidth[i].Visible = true;
                txtBoxCoefScaleY[i].Visible = true;
            }
            labelLineWidth.Visible = true;
            labelCoefScaleY.Visible = true;
            //textBox_sizeTextAxis.Visible = true;
            groupBoxAxisCentral.Visible = true;
            groupBox_outParameter.Visible = true;
            dataGridViewAnalysis.Visible = true;

            buttonNarrowZGPane.Visible = false;
        }


        // ОТОБРАЖЕНИЕ ЛЕГЕНДЫ
        private void checkBoxLegend_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxLegend.Checked) pane.Legend.IsVisible = true;
            else pane.Legend.IsVisible = false;
            // обновляем график
            zedGraphPane.Invalidate();
        }


        // МИНИМАЛЬНОЕ ЗНАЧЕНИЕ В МАССИВЕ
        private double minValueArray(double[] array)
        {
            int length = array.Length;
            minValue = array[0];
            for (int i = 1; i < length; i++)
            {
                if (array[i] < minValue) minValue = array[i];
            }

            return minValue;
        }


        // МАКСИМАЛЬНОЕ ЗНАЧЕНИЕ В МАССИВЕ
        private double maxValueArray(double[] array)
        {
            int length = array.Length;
            maxValue = array[0];
            for (int i = 1; i < length; i++)
            {
                if (array[i] > maxValue) maxValue = array[i];
            }

            return maxValue;
        }


        // СОХРАНИЕ ДАННЫХ В EXCEL-ФАЙЛ
        private void сохранитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (stringData != null)
            {
                // диалоговое окно для выбора места сохранения файла
                SaveFileDialog sfd = new SaveFileDialog();
                if (path_file_xlsx == null) sfd.InitialDirectory = Environment.CurrentDirectory;
                else
                {
                    n = path_file_xlsx.LastIndexOf("\\");
                    sfd.InitialDirectory = path_file_xlsx.Remove(n);
                }
                sfd.Filter = "Microsoft Office Excel *.xlsx|*.xlsx";
                str = DateTime.Now.ToString();
                sfd.FileName = "Simulation_Data " + str.Replace(':', '.');

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Stopwatch time = new Stopwatch();
                    time.Start();

                    Cursor = Cursors.WaitCursor;

                    // директория сохранения файла
                    path_file_xlsx = sfd.FileName;

                    // начальные приготовления для excel
                    Excel.Application app_data = null;
                    Excel.Workbook book_data = null;
                    Excel.Worksheet sheet_data = null;
                    Excel.Range range_data = null;
                    app_data = new Excel.Application();
                    app_data.Visible = false;
                    book_data = app_data.Workbooks.Add(System.Type.Missing);
                    sheet_data = (Excel.Worksheet)book_data.Worksheets.get_Item(1);



                    // записываем переменные в excel-файл                                             
                    L = stringData.Length;  // кол-во массивов
                    // ищем максимальную длину массива
                    for (int i = 0; i < L; i++)
                    {
                        n = plotData[i].Length;
                        if (i == 0) n_max = n;
                        else if (n > n_max) n_max = n;
                    }
                    // заполнение массива и запись в excel файл
                    object[,] data = new object[n_max, L];
                    for (int i = 0; i < L; i++)
                    {
                        n = plotData[i].Length;
                        sheet_data.Cells[1, i + 1] = stringData[i];
                        for (int j = 0; j < n; j++) data[j, i] = plotData[i][j];
                    }
                    range_data = sheet_data.get_Range("A2", sheet_data.Cells[n_max + 1, L]);
                    range_data.set_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault, data);



                    // закрываем и сохраняем файл excel
                    try
                    {
                        book_data.SaveAs(path_file_xlsx, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                    }
                    catch { }
                    book_data.Close(true, System.Type.Missing, System.Type.Missing);
                    app_data.Quit();
                    Cursor = Cursors.Default;
                    time.Stop();
                    MessageBox.Show("Данные успешно сохранены\n" + time.Elapsed.ToString());
                }
                else MessageBox.Show("Данные не удалось сохранить");
            }
            else MessageBox.Show("Нет данных");
        }


        // ЗАГРУЗКА ДАННЫХ
        private void загрузитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Текстовые файлы(*.txt)|*.txt";
            ofd.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                DialogResult result = MessageBox.Show("Использовать первую строку в файле в качестве заголовка?", " ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                n_max = 100000;
                string[] strPlot = new string[1];
                double[][] data = new double[1][];

                using (StreamReader sr = new StreamReader(ofd.FileName, Encoding.GetEncoding(1251)))
                {
                    n = -1;
                    string str_int = "";

                    // считываем по строчкам
                    while ((str = sr.ReadLine()) != null)
                    {
                        // если пустая строка попалась
                        if (str.Length == 0) continue;

                        // 1-ая строка считываемого файла
                        if (n == -1)
                        {
                            // считаем кол-во столбцов
                            L = str.Length - str.Replace(";", "").Length + 1;
                            if (str[str.Length - 1] == ';') L--;

                            // объявляем массивы размером, равным кол-ву столбцов
                            strPlot = new string[L];
                            data = new double[L][];
                            for (int i = 0; i < L; i++) data[i] = new double[n_max];

                            // формируем названия (заголовки) отображаемых данных
                            L = str.Length;
                            m = 0;
                            for (int i = 0; i <= L; i++)
                            {
                                if ((i != L) && (str[i] != ';')) str_int += str[i];
                                else
                                {
                                    if (result == DialogResult.Yes) strPlot[m] = str_int;
                                    else strPlot[m] = m.ToString();
                                    str_int = "";
                                    m++;
                                }
                            }

                            n = 0;
                            // если первая строка содержит заголовки, то идем на след. такт, чтобы один такт не пропускался
                            if (result == DialogResult.Yes) continue;
                        }


                        // считываем данные из файла
                        L = str.Length;
                        m = 0;
                        for (int i = 0; i <= L; i++)
                        {
                            if ((i != L) && (str[i] != ';')) str_int += str[i];
                            else
                            {
                                try
                                {
                                    data[m][n] = Convert.ToDouble(str_int);
                                    str_int = "";
                                    m++;
                                }
                                catch { }
                            }
                        }

                        n++;

                        // смотрим, если длину массива надо увеличить
                        if (n >= n_max)
                        {
                            n_max *= 2;
                            for (int i = 0; i < strPlot.Length; i++) Array.Resize(ref data[i], n_max);
                        }
                    }
                }

                // смотрим, если массив оказался длинным и надо его уменьшить
                if (n < n_max) for (int i = 0; i < strPlot.Length; i++) Array.Resize(ref data[i], n);

                // теперь отображаем загруженные данные
                Initialization(data, strPlot);
            }
        }
    }
}