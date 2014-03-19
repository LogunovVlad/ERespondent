﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using ERespondent.CheckData;
using ERespondent.UtilityFunction;
using ERespondent.Entity;

namespace ERespondent
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        private DirectionEnergySave _formDirEnSave;
        private TypeFuel _formTypeFuel;
        private OKPO _okpoForm;
        private E_RespondentDataContext _db;

        private void MainForm_Load(object sender, EventArgs e)
        {
            ScreenResolution();
            try
            {
                //заполянем комбобоксы для раздела 1 
                FillComboBox(Section1_dataGrid1ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
                FillComboBox(Section1_dataGrid2ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
                FillComboBox(Section1_dataGrid3ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");

                //заполянем комбобоксы для раздела 2
                FillComboBoxLinq(Section2_dataGrid1ColumnV, "DestinationSave");
                FillComboBoxLinq(Section2_dataGrid2ColumnV, "DestinationSave");
                FillComboBoxLinq(Section2_dataGrid3ColumnV, "DestinationSave");

                FillComboBoxLinq(Section2_dataGrid1ColumnD, "TypeFuelEnergy");
                FillComboBoxLinq(Section2_dataGrid1ColumnE, "TypeFuelEnergy");
                FillComboBoxLinq(Section2_dataGrid2ColumnD, "TypeFuelEnergy");
                FillComboBoxLinq(Section2_dataGrid2ColumnE, "TypeFuelEnergy");
                FillComboBoxLinq(Section2_dataGrid3ColumnD, "TypeFuelEnergy");
                FillComboBoxLinq(Section2_dataGrid3ColumnE, "TypeFuelEnergy");

                FillSection3Table1();
                statusStrip1.Items[0].Text = "Соединение установлено!";
            }
            catch (SqlException)
            {
                MessageBox.Show("Ошибка сервера!");
                statusStrip1.Items[0].Text = "Отсутствует подключение к базе данных!";
            }

            //для первого раздела
            Section1_dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);

            //для второго раздела
            Section2_dataGridViewHeader2_1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridViewHeader2_2.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView2.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView3.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
        }

        #region РАЗДЕЛ 1 (tab1)

        #region datagridView1 - По плану мероприятий отчетного года; datagridView2 - Дополнительные мероприятия
        /// <summary>
        /// Событие. Происходит при клике-нажатии на ячейку
        /// Выбор ячейки с в таблице(для того, что бы comboBox раскрывался сразу)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        DataGridView grid;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            grid = (DataGridView)sender;
            //если текстБокс, чтобы не вызывать DropDown. (проверка для поля итого)
            string tt = grid.CurrentCell.EditType.ToString();
            if (tt.Equals("System.Windows.Forms.DataGridViewComboBoxEditingControl"))
            {
                if (grid.Columns[e.ColumnIndex].Index == 2)
                {
                    ((ComboBox)grid.EditingControl).DroppedDown = true;
                    ((ComboBox)grid.EditingControl).SelectionChangeCommitted += new EventHandler(method);
                }
                else
                {
                    ((ComboBox)grid.EditingControl).DroppedDown = true;
                    ((ComboBox)grid.EditingControl).SelectionChangeCommitted -= new EventHandler(method);
                }
            }
        }

        /// <summary>
        /// Событие на вывод строки "Итого"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxTable1_Click(object sender, EventArgs e)
        {
            DataGridView grid = null;
            CheckBox c = ((CheckBox)sender);
            if (c.Name.Equals("checkBoxTable1"))
            {
                grid = Section1_dataGridView1;
            }
            if (c.Name.Equals("checkBoxTable2"))
            {
                grid = Section1_dataGridView2;
            }
            if (grid.RowCount > 1)
            {
                if (grid != null)
                {
                    grid.AllowUserToAddRows = false;
                    int _rowCount = grid.RowCount;
                    DataGridViewRow _newRow = new DataGridViewRow();
                    if (c.Checked)
                    {
                        for (int i = 0; i < grid.ColumnCount; i++)
                        {
                            _newRow.Cells.Add(new DataGridViewTextBoxCell());
                        }
                        _newRow.ReadOnly = true;
                        grid.Rows.InsertRange(_rowCount, _newRow);
                        _rowCount = grid.RowCount;

                        grid[2, _rowCount - 1].Value = "Итого";
                        StyleTotalCells(grid);
                        FillRowValue_X(grid, grid.RowCount - 1, 3, 5);
                        AutoTotalSumm.TotalSumm(grid, 6); //6 - потому что расчет Итого ведется начиная с 6 колонки таблицы (графа 2 - Экономия ТЭР) 
                        AutoTotalSumm.FillTotalRow(grid, 6);
                    }
                    else
                    {
                        grid.Rows.RemoveAt(_rowCount - 1);
                        grid.AllowUserToAddRows = true;
                    }
                }
            }
        }

        /// <summary>
        /// Активирует элемент управления CheckBox для добавления строки ИТОГО
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            var grid = ((DataGridView)sender);

            if (grid.Tag.Equals("T1"))
                checkBoxTable1.Enabled = true;
            else
                checkBoxTable2.Enabled = true;

            grid.Focus();
        }

        /// <summary>
        /// Вызываем пересчет ИТОГО
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1and2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var grid = (DataGridView)sender;
            grid.EndEdit();
            if (grid.Name.Equals("Section1_dataGridView1"))
            {
                AutoTotalSumm.TotalSumm(grid, 6); //6 - потому что расчет Итого ведется начиная с 6 колонки таблицы (графа 2 - Экономия ТЭР) 
                if (checkBoxTable1.Checked)
                {
                    AutoTotalSumm.FillTotalRow(grid, 6);
                }
            }
            else
            {
                if (grid.Name.Equals("Section1_dataGridView2"))
                {
                    AutoTotalSumm.TotalSumm(grid, 6); //6 - потому что расчет Итого ведется начиная с 6 колонки таблицы (графа 2 - Экономия ТЭР) 
                    if (checkBoxTable2.Checked)
                    {
                        AutoTotalSumm.FillTotalRow(grid, 6);
                    }
                }
            }
            //Пересчитать всего по разделу, если изменились значения
            if (checkBoxTable3.Checked)
            {
                AutoTotalSumm.TotalSummGrid3(Section1_dataGridView3, 6);
                AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3, 6);
                AutoTotalSumm.FillGrid3(Section1_dataGridView3, 6);
            }
        }

        /// <summary>
        /// Для изменения ширины столбцов, для каждой таблицы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            int indexCol = e.Column.Index;
            int newWidth = e.Column.Width;

            Section1_dataGridView1.Columns[indexCol].Width = newWidth;
            Section1_dataGridView2.Columns[indexCol].Width = newWidth;
            Section1_dataGridView3.Columns[indexCol].Width = newWidth;
        }
        #endregion

        #region dataGridView3
        /// <summary>
        /// Для datagridView 3 нужно указать столбца которые нельзя редактировать
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView grid = ((DataGridView)sender);
            //int _rowNumber = grid.CurrentRow.Index;
            int _rowNumber = grid.RowCount - 2;
            checkBoxTable3.Enabled = true;

            //!!!!!!!!!!!!!Пустить по циклу
            grid.Rows[_rowNumber].Cells[4].Value = "X";
            grid.Rows[_rowNumber].Cells[4].ReadOnly = true;
            grid.Rows[_rowNumber].Cells[4].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Rows[_rowNumber].Cells[4].Style.BackColor = Color.LightGray;

            grid.Rows[_rowNumber].Cells[5].Value = "X";
            grid.Rows[_rowNumber].Cells[5].ReadOnly = true;
            grid.Rows[_rowNumber].Cells[5].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Rows[_rowNumber].Cells[5].Style.BackColor = Color.LightGray;

            for (int cellInd = 7; cellInd < 15; cellInd++)
            {
                grid.Rows[_rowNumber].Cells[cellInd].Value = "X";
                grid.Rows[_rowNumber].Cells[cellInd].ReadOnly = true;
                grid.Rows[_rowNumber].Cells[cellInd].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid.Rows[_rowNumber].Cells[cellInd].Style.BackColor = Color.LightGray;
            }

            grid.ClearSelection();
            grid.Focus();
        }

        /// <summary>
        /// 3-ий checkBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxTable3_Click(object sender, EventArgs e)
        {
            DataGridView grid = null;
            grid = Section1_dataGridView3;
            CheckBox c = ((CheckBox)sender);

            if (grid != null && grid.RowCount > 1)
            {
                grid.AllowUserToAddRows = false;
                int _rowCount = grid.RowCount;
                DataGridViewRow _newRow = new DataGridViewRow();
                if (c.Checked)
                {
                    for (int i = 0; i < grid.ColumnCount; i++)
                    {
                        _newRow.Cells.Add(new DataGridViewTextBoxCell());
                    }
                    _newRow.ReadOnly = true;
                    grid.Rows.InsertRange(_rowCount, _newRow);
                    _rowCount = grid.RowCount;

                    grid[2, _rowCount - 1].Value = "Итого";
                    StyleTotalCells(grid);
                    FillRowValue_X(grid, grid.RowCount - 1, 3, 5);
                    FillRowValue_X(grid, grid.RowCount - 1, 7, 14);

                    //добавим строку ИТОГО ПО РАЗДЕЛУ 1
                    _rowCount = grid.RowCount;
                    _newRow = new DataGridViewRow();
                    for (int i = 0; i < grid.ColumnCount; i++)
                    {
                        _newRow.Cells.Add(new DataGridViewTextBoxCell());
                    }
                    grid.Rows.InsertRange(_rowCount, _newRow);
                    _rowCount = grid.RowCount;
                    grid[2, _rowCount - 1].Value = "Всего по разделу I";
                    grid[2, _rowCount - 1].ReadOnly = true;
                    grid[2, _rowCount - 1].Style.Alignment = DataGridViewContentAlignment.TopLeft;
                    grid[2, _rowCount - 1].Style.BackColor = Color.LightGray;
                    FillRowValue_X(grid, grid.RowCount - 1, 3, 5);
                    //end                   
                    AutoTotalSumm.TotalSummGrid3(Section1_dataGridView3, 6);
                    AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3, 6);
                    AutoTotalSumm.FillGrid3(grid, 6);

                }
                else
                {
                    grid.Rows.RemoveAt(_rowCount - 1);
                    grid.Rows.RemoveAt(_rowCount - 2);
                    grid.AllowUserToAddRows = true;
                }
            }
        }

        /// <summary>
        /// Пересчет итого, при редактировании ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var grid = (DataGridView)sender;
            grid.EndEdit();
            if (checkBoxTable3.Checked)
            {
                AutoTotalSumm.TotalSummGrid3(grid, 6);
                AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3, 6);
                AutoTotalSumm.FillGrid3(grid, 6);
            }
        }
        #endregion

        #region Главное меню
        private void справочникКодовОКПООрганизацийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _okpoForm = new OKPO();
            //указываем владельца
            _okpoForm.Owner = this;
            _okpoForm.Show();
        }

        /// <summary>
        /// Форма для вывода перечня основных энергосбережений
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void переченьОсновныхНаправленийЭнергосбереженияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _formDirEnSave = new DirectionEnergySave();
            _formDirEnSave.Owner = this;
            _formDirEnSave.Show();
        }

        private void видыТопливаИЭнергииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _formTypeFuel = new TypeFuel();
            _formTypeFuel.Owner = this;
            _formTypeFuel.Show();
        }

        private void соединитьСБазойДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ConnectionDB connection = new ConnectionDB();
            SqlConnection conn = connection.CreateConnection();
            if (conn != null)
            {
                statusStrip1.Items[0].Text = "Соединение установлено!";
            }
            else
            {
                statusStrip1.Items[0].Text = "Ошибка соединения!";
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Контрольные функции для раздела 1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void контрольныеФункцииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            #region Снимаем выделения с ячеек таблицы и активируем строку ИТОГО
            Section1_dataGridView1.CurrentCell = null;
            Section1_dataGridView2.CurrentCell = null;
            Section1_dataGridView3.CurrentCell = null;
            Section2_dataGridView1.CurrentCell = null;
            Section2_dataGridView2.CurrentCell = null;
            Section2_dataGridView3.CurrentCell = null;
            Section3_T3.CurrentCell = null;
            Section3_T4.CurrentCell = null;
            Section3_T5.CurrentCell = null;

            if (Section1_dataGridView1.RowCount > 1)
            { checkBoxTable1.Checked = true; }
            if (Section1_dataGridView2.RowCount > 1)
            { checkBoxTable2.Checked = true; }
            /*if (Section1_dataGridView3.RowCount > 1)           
            {*/
            Section1_dataGridView3.AllowUserToAddRows = false;
            Section1_dataGridView3.RowCount = 1;
            checkBoxTable3.Checked = true;

            if (Section2_dataGridView1.RowCount > 1)
            { Section2_checkBoxTable1.Checked = true; }
            if (Section2_dataGridView2.RowCount > 1)
            { Section2_checkBoxTable2.Checked = true; }
            if (Section2_dataGridView3.RowCount > 1)
            { Section2_checkBoxTable3.Checked = true; }
            #endregion

            CheckProtocol.ErrorForAllSection.Clear();
            ControlFunction controlObj = new ControlFunction();
            controlObj.CheckSection(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3,
                "РАЗДЕЛ 1: ВЫПОЛНЕНИЕ МЕРОПРИЯТИЙ ПО ЭКОНОМИИ ТОПЛИВНО-ЭНЕРГЕТИЧЕСКИХ РЕСУРСОВ (ТЭР)\n", 8);

            controlObj = new ControlFunction();
            controlObj.CheckSection(Section2_dataGridView1, Section2_dataGridView2, Section2_dataGridView3, "\n\nРАЗДЕЛ 2: ВЫПОЛНЕНИЕ МЕРОПРИЯТИЙ ПО" +
            "УВЕЛИЧЕНИЮ ИСПОЛЬЗОВАНИЯ МЕСТНЫХ ВИДОВ ТОПЛИВА, ОТХОДОВ ПРОИЗВОДСТВА И ДРУГИХ ВТОРИЧНЫХ И ВОЗОБНОВЛЯЕМЫХ ЭНЕРГОРЕСУРСОВ (МВТ)\n", 10);
            DataGridView[] massGrid = new DataGridView[] { Section1_dataGridView1, Section1_dataGridView2, Section2_dataGridView1, Section2_dataGridView2 };

            //Фактическое значение из Раздела 3 таблицы 3 (строка 1)
            controlObj.CheckSection3Table3Row1(massGrid, Section3_T3);

            //Фактическое значение из Раздела 3 таблицы 3 (Строка 2)
            controlObj.CheckSection3Table3Row2(Section1_dataGridView1, Section1_dataGridView2, Section3_T3);
            
            //Фактическое значение из Раздела 3 таблицы 3 (Строка 3)
            controlObj.CheckSection3Table3Row3(Section2_dataGridView1, Section2_dataGridView2, Section3_T3);

            //Фактическое значение из Раздела 3 таблицы 3 (Строка 4)
            controlObj.CheckSection3Table3Row4(Section1_dataGridView3, Section2_dataGridView3, Section3_T3);

            controlObj.ShowListError();

        }
        #endregion

        #region Работа с ячейками таблицы
        /// <summary>
        /// Обработчик выбора значения в комбобокс
        /// и подстановки в 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void method(object sender, EventArgs e)
        {
            SqlConnection _conn = new ConnectionDB().CreateConnection();
            SqlDataAdapter _daMain;

            using (_conn)
            {
                string str = null;
                string whereStr = null;
                string queryString = null;
                bool flag = true;
                int _rowIndex = 0;
                try
                {
                    str = ((ComboBox)grid.EditingControl).SelectedValue.ToString();
                    //узнаем индекс строки, чтобы знать куда подставлять значения их выборки(подставляем код направления и единицы измерения)
                    _rowIndex = grid.CurrentCell.RowIndex;
                }
                catch (NullReferenceException)
                {
                    grid.CancelEdit();
                    flag = false;
                    return;
                }
                DataTable tb = null;
                if (flag)
                {
                    //берем исходя из выбранного элемента в ComboBox код направления(CodeDirection) и...
                    whereStr = String.Format("where CodeRecord='{0}'", str);
                    //...делаем выборкув соответствии с этим условием
                    queryString = String.Format("select CodeDirection, Unit, DestinationsSave from DestinationSave " + " {0}", whereStr);
                    //                   
                    //АДСКИЙ TRY CATCH! Из-за истинно арийских методов реализации редактирования элемента ColumnComboBox!
                    //З.Ы. При выборе элемента в комбобокс(в колонке комбобоксов), если переключиться на другую ячейку (не комбобокс),
                    //например текстовую или календарь, то большой брат дает добро жить без исключений. Если выбрали сразу же другую ячейку,
                    //в которой лежит комбобокс -> тьма! Т.к. по непонятным причинам, не происходит (точнее оно происходит, но возвращает
                    //не коллекцию Row) какое-то событие, которое достает вот отсюда << ((ComboBox)grid.EditingControl).SelectedValue >>
                    //код записи, по которому мы достаем данные для подстановки.                    
                    try
                    {
                        ///Переделать выборку (Один раз загружаем в DataTable и с ней работаем)
                        ///Сейчас каждый раз обращаемся к базе!((
                        SqlCommand comm = new SqlCommand(queryString, _conn);
                        tb = new DataTable();
                        _daMain = new SqlDataAdapter(comm);
                        _daMain.Fill(tb);
                    }
                    catch (SqlException)
                    {
                        //туoo возвращаемся и все работает
                        return;
                    }
                }
                try
                {
                    grid.Rows[_rowIndex].Cells[0].Value = tb.Rows[0][0].ToString();
                    //чтобы подстановка не попадала в строку "Итого по разделу"
                    if (!grid.Tag.Equals("Пункт \"По мероприятиям предшествующего года внедрения\""))
                    {
                        grid.Rows[_rowIndex].Cells[4].Value = tb.Rows[0][1].ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Проверьте строку\n" + ex.StackTrace);
                }
                //grid.SendToBack();                       
            }
        }

        /// <summary>
        /// Подстановка значений из таблиц в ComboBox
        /// </summary>
        /// <param name="col">Столбец с типом comboBox из таблицы</param>
        /// <param name="queryString">Строка с sql-запросом для заполнения SqlDataAdapter</param>
        /// <param name="displayMember">Поле, которое будет отображаться в выпадающем списке</param>
        /// <param name="valueMember">Первичный ключ</param>
        private void FillComboBox(DataGridViewComboBoxColumn col, string queryString, string displayMember, string valueMember)
        {
            SqlConnection _conn = new ConnectionDB().CreateConnection();
            SqlDataAdapter _daMain;
            using (_conn)
            {
                SqlCommand comm = new SqlCommand(queryString, _conn);
                DataTable tb = new DataTable();
                _daMain = new SqlDataAdapter(comm);
                _daMain.Fill(tb);
                col.DataSource = tb;
                col.DisplayMember = displayMember;
                col.ValueMember = valueMember;
                col.DropDownWidth = 1200;
            }
        }

        /// <summary>
        /// Метод заполняющий ячейки
        /// </summary>
        /// <param name="grid">Таблица, в которой заполняются ячейки</param>
        /// <param name="_rowCount">Количество строк в таблице</param>
        /// <param name="from">От какой ячейки заполнять</param>
        /// <param name="to">До какой ячейки заполнять</param>
        private static void FillRowValue_X(DataGridView grid, int _numberRow, int from, int to)
        {
            for (int colInd = from; colInd < to + 1; colInd++)
            {
                grid[colInd, _numberRow].Value = "X";
                grid[colInd, _numberRow].ReadOnly = true;
                grid[colInd, _numberRow].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid[colInd, _numberRow].Style.BackColor = Color.LightGray;
            }
        }
        #endregion

        #region Контекстное меню таблицы 1 и 2 !!!!!!!!Переделать

        private DataGridView gridContext = null;
        private ContextMenuStrip strip;
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            gridContext = ((DataGridView)sender);

            ToolStripMenuItem delete = new ToolStripMenuItem();
            ToolStripMenuItem add = new ToolStripMenuItem();
            ToolStripMenuItem copyText = new ToolStripMenuItem();
            ToolStripMenuItem pasteText = new ToolStripMenuItem();
            ToolStripMenuItem cutText = new ToolStripMenuItem();

            if (strip == null)
            {
                strip = new ContextMenuStrip();
                delete.Text = "Удалить строку";
                delete.Name = "delete";
                add.Text = "Добавить строку";
                add.Name = "add";
                copyText.Text = "Копировать текст ячейки";
                copyText.Name = "copy";
                pasteText.Text = "Вставить";
                pasteText.Name = "paste";
                cutText.Text = "Вырезать";
                cutText.Name = "cut";
                strip.Items.Add(delete);
                strip.Items.Add(add);
                //strip.Items.Add(copyText);
                //strip.Items.Add(pasteText);
                //strip.Items.Add(cutText);

                strip.Items["delete"].Click += new EventHandler(delete_Click);
                strip.Items["add"].Click += new EventHandler(add_Click);
                //strip.Items["copy"].Click += new EventHandler(copyText_Click);
                //strip.Items["paste"].Click += new EventHandler(pasteText_Click);

            }
            //чтобы нельзя было удалить новую пустую строку
            if (gridContext.CurrentRow.Index == gridContext.RowCount - 1)
            {
                strip.Items["delete"].Enabled = false;
            }
            else { strip.Items["delete"].Enabled = true; }


            e.Control.ContextMenuStrip = strip;
        }

        private void delete_Click(object sender, EventArgs e)
        {
            gridContext.Rows.RemoveAt(gridContext.CurrentRow.Index);
        }

        private void add_Click(object sender, EventArgs e)
        {
            gridContext.Rows.Add();
        }

        private void copyText_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(gridContext.SelectedCells[0].Value.ToString());
        }

        private void pasteText_Click(object sender, EventArgs e)
        {
            gridContext.SelectedCells[0].Value = Clipboard.GetText();
            gridContext.UpdateCellValue(6, 0);

        }
        #endregion

        #endregion


        #region РАЗДЕЛ 2 (tab2)

        #region Таблица 1 и 2 и 3
        /// <summary>
        /// Метод для заполнения comboBox для 1 и 2 таблицы (раздел 2)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FillComboBoxLinq(DataGridViewComboBoxColumn boxColumn, string nameTable)
        {
            /*BindingSource bs = new BindingSource();
            bs.DataSource = packFill;*/
            _db = new E_RespondentDataContext();

            switch (nameTable)
            {
                case "DestinationSave":
                    var save = from c in _db.DestinationSave
                               select c;
                    boxColumn.DataSource = save;
                    boxColumn.DisplayMember = "DestinationsSave";
                    boxColumn.ValueMember = "Coderecord";
                    break;
                case "TypeFuelEnergy":
                    var energy = from c in _db.TypeFuelEnergy
                                 select c;
                    boxColumn.DataSource = energy;
                    boxColumn.DisplayMember = "CodeTypeFuel";
                    boxColumn.ValueMember = "CodeRecord";
                    break;
            }
            boxColumn.DropDownWidth = 1200;
        }

        private DataGridView _gridSection2;
        /// <summary>
        /// Событие, которое происходит по нажатию на ячейку dataGridView
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if ("System.Windows.Forms.DataGridViewComboBoxEditingControl".Equals(_gridSection2.CurrentCell.EditType.ToString()))
            {
                if (e.ColumnIndex == 2)
                {
                    ((ComboBox)_gridSection2.EditingControl).DroppedDown = true;
                    ((ComboBox)_gridSection2.EditingControl).SelectionChangeCommitted += new EventHandler(Section2_SelectedIndexChanged);
                }
                else
                {
                    ((ComboBox)_gridSection2.EditingControl).DroppedDown = true;
                    ((ComboBox)_gridSection2.EditingControl).SelectionChangeCommitted -= new EventHandler(Section2_SelectedIndexChanged);
                }
            }
        }

        /// <summary>
        /// Подстановка "Код" и "Единицы измерения" при выборе из ComboBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_SelectedIndexChanged(object sender, EventArgs e)
        {
            _db = new E_RespondentDataContext();
            int index = Convert.ToInt32(((ComboBox)sender).SelectedValue);
            DestinationSave row = (from c in _db.DestinationSave
                                   where c.CodeRecord == index
                                   select c).Single<DestinationSave>();
            _gridSection2.CurrentRow.Cells[0].Value = row.CodeDirection;
            if (!_gridSection2.Tag.Equals("T3"))
            {
                _gridSection2.CurrentRow.Cells[6].Value = row.Unit;
            }
            _gridSection2.EndEdit();
        }
        #endregion

        #region Section2_checkBoxTable
        /// <summary>
        /// Событие для обработки нажатия CheckBox (Для добавления строки итого)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_checkBoxTable_Click(object sender, EventArgs e)
        {
            switch (((CheckBox)sender).Tag.ToString())
            {
                case "Section2_chB1":
                    Activate_1and2_CheckBox(sender, Section2_dataGridView1);
                    break;
                case "Section2_chB2":
                    Activate_1and2_CheckBox(sender, Section2_dataGridView2);
                    break;
                case "Section2_chB3":
                    Activate_3_CheckBox(sender, Section2_dataGridView3);
                    break;
            }
        }

        /// <summary>
        /// Для добавления строки итого для 1 и 2 таблицы
        /// </summary>
        /// <param name="sender"></param>
        private void Activate_1and2_CheckBox(object sender, DataGridView grid)
        {
            //когда нажата, добавляем строку
            if (((CheckBox)sender).Checked)
            {
                InsertTextRow(grid);
                grid[2, grid.RowCount - 1].Value = "Итого";
                grid[2, grid.RowCount - 1].Selected = true;
                StyleTotalCells(grid);
                FillRowValue_X(grid, grid.RowCount - 1, 3, 7);
                AutoTotalSumm.TotalSumm(grid, 8);
                AutoTotalSumm.FillTotalRow(grid, 8);
            }
            //когда отжата - удаляем строку итого
            else
            {
                if (grid.Rows.Count > 0)
                {
                    grid.Rows.RemoveAt(grid.Rows.Count - 1);
                    grid.AllowUserToAddRows = true;
                }
            }
        }

        /// <summary>
        /// Для добавления строки "Итого" и "Всего по разделу" для 3 талблицы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="grid"></param>
        private void Activate_3_CheckBox(object sender, DataGridView grid)
        {
            if (((CheckBox)sender).Checked)
            {
                InsertTextRow(grid);
                grid[2, grid.RowCount - 1].Value = "Итого";
                StyleTotalCells(grid);
                FillRowValue_X(grid, grid.RowCount - 1, 3, 7);
                FillRowValue_X(grid, grid.RowCount - 1, 9, 16);

                InsertTextRow(grid);
                grid[2, grid.RowCount - 1].Value = "Всего по разделу 2";
                StyleTotalCells(grid);
                FillRowValue_X(grid, grid.RowCount - 1, 3, 7);

                AutoTotalSumm.TotalSummGrid3(grid, 8);
                AutoTotalSumm.TotalAll1Section(Section2_dataGridView1, Section2_dataGridView2, Section2_dataGridView3, 8);
                AutoTotalSumm.FillGrid3(grid, 8);

                /*AutoTotalSumm.TotalSummGrid3(Section1_dataGridView3, 6);
                AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3, 6);
                AutoTotalSumm.FillGrid3(Section1_dataGridView3, 6);*/
            }
            else
            {
                grid.Rows.RemoveAt(grid.Rows.Count - 1);
                grid.Rows.RemoveAt(grid.Rows.Count - 1);
                grid.AllowUserToAddRows = true;
            }

        }

        /// <summary>
        /// Изменение стиля ячеек для строки Итого
        /// </summary>
        /// <param name="grid"></param>
        private static void StyleTotalCells(DataGridView grid)
        {
            grid[2, grid.RowCount - 1].ReadOnly = true;
            grid[2, grid.RowCount - 1].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            grid[2, grid.RowCount - 1].Style.BackColor = Color.LightGray;
        }

        #endregion

        #region Определяем текущую таблицу
        private void Section2_dataGridView1_Click(object sender, EventArgs e)
        {
            _gridSection2 = Section2_dataGridView1;
        }

        private void Section2_dataGridView2_Click(object sender, EventArgs e)
        {
            _gridSection2 = Section2_dataGridView2;
        }

        private void Section2_dataGridView3_Click(object sender, EventArgs e)
        {
            _gridSection2 = Section2_dataGridView3;
        }
        #endregion

        /// <summary>
        /// Собатие, которое происходит при изменении размера столбца,
        /// при этом все связанные с этой колонкой столбцы тоже меняют размер
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewSection2_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            int indexCol = e.Column.Index;
            int newWidth = e.Column.Width;

            Section2_dataGridViewHeader2_1.Columns[indexCol].Width = newWidth;
            Section2_dataGridViewHeader2_2.Columns[indexCol].Width = newWidth;
            Section2_dataGridView1.Columns[indexCol].Width = newWidth;
            Section2_dataGridView2.Columns[indexCol].Width = newWidth;
            Section2_dataGridView3.Columns[indexCol].Width = newWidth;
        }

        /// <summary>
        /// Активировать элемент управления (checkBox) для Итого
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_dataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            switch (((DataGridView)sender).Tag.ToString())
            {
                case "T1":
                    Section2_checkBoxTable1.Enabled = true;
                    break;
                case "T2":
                    Section2_checkBoxTable2.Enabled = true;
                    break;
                case "T3":
                    FillRowValue_X(Section2_dataGridView3, Section2_dataGridView3.RowCount - 2, 6, 7);
                    FillRowValue_X(Section2_dataGridView3, Section2_dataGridView3.RowCount - 2, 9, 16);
                    Section2_checkBoxTable3.Enabled = true;
                    break;
            }
            // ((DataGridView)sender).ClearSelection();
        }

        /// <summary>
        /// Создает строку c текстовыми полями и вставляет в конец (для поля "Итого" и "Всего по разделу")
        /// </summary>
        /// <param name="grid"></param>
        private void InsertTextRow(DataGridView grid)
        {
            grid.AllowUserToAddRows = false;
            DataGridViewRow _newRow = new DataGridViewRow();
            for (int i = 0; i < Section2_dataGridView1.ColumnCount; i++)
            {
                _newRow.Cells.Add(new DataGridViewTextBoxCell());
            }
            _newRow.ReadOnly = true;
            grid.Rows.InsertRange(grid.Rows.Count, _newRow);
        }

        /// <summary>
        /// Пересчет значений ИТОГО при изменении значения ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            switch (((DataGridView)sender).Name)
            {
                case "Section2_dataGridView1":
                    AutoTotalSumm.TotalSumm(((DataGridView)sender), 8);
                    if (Section2_checkBoxTable1.Checked)
                    {
                        AutoTotalSumm.FillTotalRow(((DataGridView)sender), 8);
                    }
                    break;
                case "Section2_dataGridView2":
                    AutoTotalSumm.TotalSumm(((DataGridView)sender), 8);
                    if (Section2_checkBoxTable2.Checked)
                    {
                        AutoTotalSumm.FillTotalRow(((DataGridView)sender), 8);
                    }
                    break;
            }
            if (Section2_checkBoxTable3.Checked)
            {
                AutoTotalSumm.TotalSummGrid3(Section2_dataGridView3, 8);
                AutoTotalSumm.TotalAll1Section(Section2_dataGridView1, Section2_dataGridView2, Section2_dataGridView3, 8);
                AutoTotalSumm.FillGrid3(Section2_dataGridView3, 8);
            }
        }
        #endregion

        #region Раздел 3 (tab3)


        /// <summary>
        /// Заполнение таблицы 3 (Раздел 3)
        /// </summary>
        private void FillSection3Table1()
        {
            #region Таблица 3
            Section3_T3.Rows.Add(5);
            Section3_T3[0, 0].Value = "А";
            Section3_T3[1, 0].Value = "Б";
            Section3_T3[2, 0].Value = "В";
            Section3_T3[3, 0].Value = "1";
            Section3_T3[4, 0].Value = "2";
            Section3_T3[5, 0].Value = "3";

            for (int i = 0; i < 6; i++)
            {
                Section3_T3[i, 0].ReadOnly = true;
                Section3_T3[i, 0].Style.BackColor = Color.FromArgb(181, 181, 181);
                Section3_T3[i, 0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            Section3_T3[0, 1].Value = "Количество мероприятий";
            Section3_T3[1, 1].Value = "1";
            Section3_T3[2, 1].Value = "ед.";

            Section3_T3[0, 2].Value = "Экономия ТЭР";
            Section3_T3[1, 2].Value = "2";
            Section3_T3[2, 2].Value = "т усл. топл.";

            Section3_T3[0, 3].Value = "Увеличение использования МВТ";
            Section3_T3[1, 3].Value = "3";
            Section3_T3[2, 3].Value = "т усл. топл.";

            Section3_T3[0, 4].Value = "Затраты на внедрение  мероприятий";
            Section3_T3[1, 4].Value = "4";
            Section3_T3[2, 4].Value = "млн.руб.";

            for (int i = 0; i < 3; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    Section3_T3[i, j].ReadOnly = true;
                    Section3_T3[i, j].Style.BackColor = Color.FromArgb(240, 240, 240);
                    Section3_T3[i, j].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
            #endregion

            #region Таблица 4
            Section3_T4.Rows.Add();
            Section3_T4[0, 0].Value = "A";
            Section3_T4[1, 0].Value = "1";
            Section3_T4[2, 0].Value = "2";
            for (int i = 0; i < 3; i++)
            {
                Section3_T4[i, 0].ReadOnly = true;
                Section3_T4[i, 0].Style.BackColor = Color.FromArgb(181, 181, 181);
                Section3_T4[i, 0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            Section3_T4.Rows.Add();
            Section3_T4[0, 1].Value = "5";
            Section3_T4[0, 1].ReadOnly = true;
            Section3_T4[0, 1].Style.BackColor = Color.FromArgb(240, 240, 240);
            Section3_T4[0, 1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            #endregion

            #region Таблица 5
            Section3_T5.Rows.Add();
            Section3_T5[0, 0].Value = "А";
            Section3_T5[1, 0].Value = "Б";
            Section3_T5[2, 0].Value = "В";

            for (int i = 0; i < 3; i++)
            {
                Section3_T5[i, 0].ReadOnly = true;
                Section3_T5[i, 0].Style.BackColor = Color.FromArgb(181, 181, 181);
                Section3_T5[i, 0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            #endregion
        }



        #endregion



        /// <summary>
        /// Устанавливает размер формы по размерам разрешения экрана
        /// </summary>
        private void ScreenResolution()
        {
            int heightScreen = Screen.PrimaryScreen.WorkingArea.Height;
            int widthScreen = Screen.PrimaryScreen.WorkingArea.Width;
            this.Location = new Point(0, 0);
            this.Height = heightScreen;
            this.Width = 1380;// widthScreen;
        }

        /// <summary>
        /// Проверка правильности ввода
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section3_T3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (((DataGridView)sender).CurrentCell.Value != null)
                {
                    double.Parse(((DataGridView)sender).CurrentCell.Value.ToString());
                }
            }
            catch (FormatException)
            {
                MessageBox.Show("Ошибка формата: введите число!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                ((DataGridView)sender).CurrentCell.Value = null;
            }

        }



    }
}

