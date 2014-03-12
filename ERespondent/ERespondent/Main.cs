using System;
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
            //определяем разрешение экрана
            ScreenResolution();

            //заполянем первый комбобоксы для первого раздела            
            try
            {
                FillComboBox(Section1_dataGrid1ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
                FillComboBox(Section1_dataGrid2ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
                FillComboBox(Section1_dataGrid3ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
            }
            catch (SqlException)
            { MessageBox.Show("Ошибка сервера!"); }
            //заполянем первый комбобоксы для второго раздела 

            /*четные нечетные строки разными цветами
             * dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightBlue;//Color.Bisque;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightCyan;//Color.Beige;*/
            // dataGridView1.Rows.Add();  

            //для первого раздела
            Section1_dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);

            //для второго раздела
            Section2_dataGridViewHeader2_1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridViewHeader2_2.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView2.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);
            Section2_dataGridView3.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridViewSection2_ColumnWidthChanged);           
        }

        /// <summary>
        /// Устанавливает размер формы по размерам разрешения экрана
        /// </summary>
        private void ScreenResolution()
        {
            int heightScreen = Screen.PrimaryScreen.WorkingArea.Height;
            int widthScreen = Screen.PrimaryScreen.WorkingArea.Width;
            this.Location = new Point(0, 0);
            this.Height = heightScreen;// 728;
            this.Width = 1380;// widthScreen;
        }

        #region РАЗДЕЛ 1(tab1)

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
                    // grid.BeginEdit(true);
                    ((ComboBox)grid.EditingControl).DroppedDown = true;
                    ((ComboBox)grid.EditingControl).SelectedIndexChanged += new EventHandler(method);
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
                        grid.Rows.InsertRange(_rowCount, _newRow);
                        _rowCount = grid.RowCount;

                        grid[2, _rowCount - 1].Value = "Итого";
                        grid[2, _rowCount - 1].ReadOnly = true;
                        grid[2, _rowCount - 1].Style.Alignment = DataGridViewContentAlignment.TopRight;
                        FillTextBoxX(grid, _rowCount, 3, 6);

                        //сумму Итого по подразделу вычисляем автоматически
                        AutoTotalSumm.TotalSumm(grid);
                        AutoTotalSumm.FillTotalRow(grid);
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
        /// Отменяет выделение текущей ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            var grid = ((DataGridView)sender);

            if (grid.Tag.Equals("Пункт \"По плану мероприятий отчетного года\""))
                checkBoxTable1.Enabled = true;
            else
                checkBoxTable2.Enabled = true;

            grid.Focus();
        }

        /// <summary>
        /// Вызываем пересчет итого
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1and2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var grid = (DataGridView)sender;
            grid.EndEdit();
            if (grid.Name.Equals("Section1_dataGridView1"))
            {
                AutoTotalSumm.TotalSumm(grid); //1 - потому что строка итого является последней в таблице 1  
                if (checkBoxTable1.Checked)
                {
                    AutoTotalSumm.FillTotalRow(grid);
                }
            }
            else
            {
                if (grid.Name.Equals("Section1_dataGridView2"))
                {
                    AutoTotalSumm.TotalSumm(grid); //1 - потому что строка итого является последней в таблице 2   
                    if (checkBoxTable2.Checked)
                    {
                        AutoTotalSumm.FillTotalRow(grid);
                    }
                }
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
            int _rowNumber = grid.CurrentRow.Index;
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

            grid.CurrentCell.Selected = false;
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
                    grid.Rows.InsertRange(_rowCount, _newRow);
                    _rowCount = grid.RowCount;
                    grid[2, _rowCount - 1].Value = "Итого";
                    grid[2, _rowCount - 1].ReadOnly = true;
                    grid[2, _rowCount - 1].Style.Alignment = DataGridViewContentAlignment.TopRight;

                    FillTextBoxX(grid, _rowCount, 3, 6);
                    FillTextBoxX(grid, _rowCount, 7, 15);

                    //добавим строку итого по разделу
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
                    FillTextBoxX(grid, _rowCount, 3, 6);
                    //end                   

                    AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3);
                    AutoTotalSumm.FillGrid3(grid);
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
                AutoTotalSumm.TotalSummGrid3(grid);
                AutoTotalSumm.TotalAll1Section(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3);
                AutoTotalSumm.FillGrid3(grid);
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
            string result = null;
            try
            {
                //5result = connection.CreateConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            statusStrip1.Items[0].Text = result;
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
            Section1_dataGridView1.EndEdit();
            Section1_dataGridView2.EndEdit();
            Section1_dataGridView3.EndEdit();

            ControlFunction controlObj = new ControlFunction();
            controlObj.CheckTable(Section1_dataGridView1);
            controlObj.CheckTable(Section1_dataGridView2);
            controlObj.CheckTotalForSection1(Section1_dataGridView1, Section1_dataGridView2, Section1_dataGridView3);
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
        private static void FillTextBoxX(DataGridView grid, int _rowCount, int from, int to)
        {
            for (int colInd = from; colInd < to; colInd++)
            {
                grid[colInd, _rowCount - 1].Value = "X";
                grid[colInd, _rowCount - 1].ReadOnly = true;
                grid[colInd, _rowCount - 1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid[colInd, _rowCount - 1].Style.BackColor = Color.LightGray;
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


        #region РАЗДЕЛ 2(tab2)
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

        #region Таблица 1
        /// <summary>
        /// Событие происходит когда выбрали вторую кладку (РАЗДЕЛ 2)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (((TabControl)sender).SelectedIndex == 1)
            {
                /*BindingSource bs = new BindingSource();
                bs.DataSource = packFill;*/
                _db = new E_RespondentDataContext();
                var packFill = from c in _db.DestinationSave
                               select c;
                Section2_dataGrid1ColumnV.DataSource = packFill;
                Section2_dataGrid1ColumnV.DisplayMember = "DestinationsSave";
                Section2_dataGrid1ColumnV.ValueMember = "Coderecord";
                Section2_dataGrid1ColumnV.DropDownWidth = 1200;

            }
        }

        private DataGridView _gridSection2;
        /// <summary>
        /// Событие, которое происходит по нажатию на ячейку dataGridView
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Section2_dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            _gridSection2 = ((DataGridView)sender);
            if (e.ColumnIndex == 2)
            {
                ((ComboBox)_gridSection2.EditingControl).DroppedDown = true;
                ((ComboBox)_gridSection2.EditingControl).SelectionChangeCommitted += new EventHandler(Section2_SelectedIndexChanged);
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
            _gridSection2.CurrentRow.Cells[6].Value = row.Unit;
        }
        #endregion

        #endregion



        /* /// <summary>
       /// Проверка правильности ввода в ячейки
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
       private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
       {
           string g=e.FormattedValue.GetType().ToString();
           if (!e.FormattedValue.GetType().ToString().Equals("System.String"))
           {
               e.Cancel = true;
               MessageBox.Show("Ошибка! Ввод отменен!");
           }
       }*/



    }
}


