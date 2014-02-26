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
            ScreenResolution();
            //заполянем первый комбобокс
            FillComboBox(dataGrid1ColumnA, "SELECT CodeDirection, CodeRecord from DestinationSave", "CodeDirection", "CodeRecord");
            FillComboBox(dataGrid1ColumnV, "SELECT DestinationsSave, CodeRecord from DestinationSave", "DestinationsSave", "Coderecord");
        }

        /// <summary>
        /// Устанавливает размер формы по размерам разрешения экрана
        /// </summary>
        private void ScreenResolution()
        {
            int heightScreen = Screen.PrimaryScreen.WorkingArea.Height;
            int widthScreen = Screen.PrimaryScreen.WorkingArea.Width;
            this.Location = new Point(0, 0);
            this.Height = heightScreen;
            this.Width = widthScreen;
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

            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            //ВЫДРАТЬ ОТСЮДА СОБЫТИЕ НА ИЗМЕНЕНИЕ ИНДЕКСА
            //COMBOBOX ДЛЯ ПОДСТАНОВКИ В ДРУГОЕ ПОЛЕ И "КОМБОБОКС"
            if (grid.Columns[e.ColumnIndex].Name.Equals("dataGrid1ColumnA"))
            {
                grid.BeginEdit(true);
                ((ComboBox)grid.EditingControl).DroppedDown = true;
                ((ComboBox)grid.EditingControl).SelectedIndexChanged += new EventHandler(method);
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
            //DataGridView grid = (DataGridView)sender;
            //получаем индекс строки, чтобы получить индекс ячейки
            int _rowIndex = grid.CurrentCell.RowIndex;
            grid.Rows[_rowIndex].Cells["dataGrid1ColumnD"].Value = "sadfsad";

            SqlConnection _conn = new ConnectionDB().CreateConnection();
            SqlDataAdapter _daMain;
            using (_conn)
            {
                //берем исходя из выбранного элемента в ComboBox код направления(CodeDirection) и...
                string whereStr = String.Format("where CodeDirection='{0}'", ((ComboBox)grid.EditingControl).SelectedValue.ToString());
                //...делаем выборкув соответствии с этим условием
                string queryString = String.Format("select DestinationsSave from DestinationSave " + " {0}", whereStr);

                //SqlCommand comm = new SqlCommand("SELECT DestinationsSave from DestinationSave WHERE CodeDirection =" + grid.CurrentCell.Value + "", _conn);                               
                SqlCommand comm = new SqlCommand(queryString, _conn);
                DataTable tb = new DataTable();
                _daMain = new SqlDataAdapter(comm);
                _daMain.Fill(tb);
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
            }
        }
    }
}
