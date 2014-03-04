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

        private void MainForm_Load(object sender, EventArgs e)
        {
            //определяем разрешение экрана
            ScreenResolution();
            //заполянем первый комбобокс            
            FillComboBox(dataGrid1ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
            FillComboBox(dataGrid2ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
            FillComboBox(dataGrid3ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");


            /*четные нечетные строки разными цветами
             * dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightBlue;//Color.Bisque;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightCyan;//Color.Beige;*/
            // dataGridView1.Rows.Add();       
        }

        /// <summary>
        /// Устанавливает размер формы по размерам разрешения экрана
        /// </summary>
        private void ScreenResolution()
        {
            int heightScreen = Screen.PrimaryScreen.WorkingArea.Height;
            int widthScreen = Screen.PrimaryScreen.WorkingArea.Width;
            this.Location = new Point(0, 0);
            this.Height = 728;// heightScreen;
            this.Width = 1380;// widthScreen;
        }

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
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                //ВЫДРАТЬ ОТСЮДА СОБЫТИЕ НА ИЗМЕНЕНИЕ ИНДЕКСА
                //COMBOBOX ДЛЯ ПОДСТАНОВКИ В ДРУГОЕ ПОЛЕ И "КОМБОБОКС"
                //if (grid.Columns[e.ColumnIndex].Name.Equals("dataGridColumnV"))
                if (grid.Columns[e.ColumnIndex].Index == 2)
                {
                    // grid.BeginEdit(true);
                    ((ComboBox)grid.EditingControl).DroppedDown = true;
                    ((ComboBox)grid.EditingControl).SelectedIndexChanged += new EventHandler(method);
                }
            }
        }

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
                catch (NullReferenceException ex)
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
                    catch (SqlException ex)
                    {
                        //тупо возвращаемся и все работает
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
                    MessageBox.Show("Проверьте строку");
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
        /// Проверка правильности ввода в ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (!e.FormattedValue.GetType().ToString().Equals("System.String"))
            {
                e.Cancel = true;
                MessageBox.Show("Ошибка! Ввод отменен!");
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
                grid = dataGridView1;
            }
            if (c.Name.Equals("checkBoxTable2"))
            {
                grid = dataGridView2;
            }

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
                }
                else
                {
                    grid.Rows.RemoveAt(_rowCount - 1);
                    grid.AllowUserToAddRows = true;
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
            DataGridView grid = ((DataGridView)sender);
            grid.CurrentCell.Selected = false;
            grid.Focus();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void контрольныеФункцииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            dataGridView2.EndEdit();
            dataGridView3.EndEdit();

            if (dataGridView1.RowCount < 2 || dataGridView2.RowCount < 2 || dataGridView3.RowCount < 2)
            {
                MessageBox.Show("Заполните все подразделы!");
            }
            else
            {

                ControlFunction controlObj = new ControlFunction();
                controlObj.CheckTable(dataGridView1);
                controlObj.CheckTable(dataGridView2);
                controlObj.CheckTotalForSection1(dataGridView1, dataGridView2, dataGridView3);

                controlObj.ShowListError();
            }
        }

        ///
        /// dataGridView3
        /// 
        #region
        /// <summary>
        /// Для datagridView 3 нужно указать столбца которые нельзя редактировать
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView grid = ((DataGridView)sender);
            int _rowNumber = grid.CurrentRow.Index;

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
                grid.Rows[_rowNumber].Cells[cellInd].ReadOnly = false;
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
            grid = dataGridView3;
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
    }
}

