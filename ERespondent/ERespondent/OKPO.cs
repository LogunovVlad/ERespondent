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
    public partial class OKPO : Form
    {
        public OKPO()
        {
            InitializeComponent();
        }
        private AddEditOKPO _addOKPO;
        private MainForm _mainForm;

        private void OKPO_Load(object sender, EventArgs e)
        {
            //определяем владельца
            _mainForm = this.Owner as MainForm;

            try
            {                                                         
                _daOKPO.Fill(_dsOKPO);                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка соединения", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //обновление записей             
            #region
            /*
            SqlCommand commandUpdate = new SqlCommand("UPDATE OKPO SET CodeOKPO=@CodeOKPO, NameOrganization=@NameOrganization " +
            "WHERE (CodeRecord=@CodeRecord)");
            commandUpdate.Connection = sqlConnection1;
            daOKPO.UpdateCommand = commandUpdate;
            SqlParameter parametr;
            //Параметр CodeOKPO
            parametr = new SqlParameter("@CodeOKPO", SqlDbType.Float);
            parametr.SourceColumn = "CodeOKPO";
            commandUpdate.Parameters.Add(parametr);
            //Параметр NameOrganization
            parametr = new SqlParameter("@NameOrganization", SqlDbType.Text);
            parametr.SourceColumn = "NameOrganization";
            commandUpdate.Parameters.Add(parametr);
            parametr = new SqlParameter("@CodeRecord", SqlDbType.Int);
            parametr.SourceColumn = "CodeRecord";            
            commandUpdate.Parameters.Add(parametr);
            #endregion
        
            // удаление записей
            #region
            SqlCommand commandDelete = new SqlCommand("DELETE FROM OKPO Where CodeRecord=@CodeRecord");
            commandDelete.Connection = sqlConnection1;
            daOKPO.DeleteCommand = commandDelete;
            SqlParameter parametrDel;
            parametrDel = new SqlParameter("@CodeOKPO", SqlDbType.Float);
            parametrDel.SourceColumn = "CodeOKPO";
            commandDelete.Parameters.Add(parametrDel);
            parametrDel = new SqlParameter("NameOrganization", SqlDbType.Text);
            parametrDel.SourceColumn = "NameOrganization";
            commandDelete.Parameters.Add(parametrDel);
            parametrDel = new SqlParameter("@CodeRecord", SqlDbType.Int);
            parametrDel.SourceColumn = "CodeRecord";
            commandDelete.Parameters.Add(parametrDel);
            */
            #endregion

        }

        /// <summary>
        /// Обработчик кнопки добавление. Вызывает новую форму, указывая, что является для нее родителем
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAdd_Click(object sender, EventArgs e)
        {
            this.Tag = "ADD";
            //вызов окна добавления
            _addOKPO = new AddEditOKPO();
            _addOKPO.Owner = this;
            _addOKPO.Show();
        }

        /// <summary>
        /// Обработчик события кнопки "Редактирование"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEdit_Click(object sender, EventArgs e)
        {
            this.Tag = "EDIT";
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Не выбрана запись для редактирования!", "Ошибка редактирования!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                _addOKPO = new AddEditOKPO();
                _addOKPO.Owner = this;
                _addOKPO.Show();
                dataGridView1.CurrentCell.Selected = true;
            }
        }

        /// <summary>
        /// Обработчик события кнопки "Удалить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btrDel_Click(object sender, EventArgs e)
        {
            try
            {
                //берем индекс выделенной ячейки и удаляем запись
                int indexRow = dataGridView1.CurrentCell.RowIndex;
                _dsOKPO.Tables["OKPO"].Rows[indexRow].Delete();

                _daOKPO.Update(_dsOKPO);
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Нет записей для удаления!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        /// <summary>
        /// Закрыть форму
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseOkpo_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OKPO_FormClosed(object sender, FormClosedEventArgs e)
        {
            _connectionOKPO.Close();
        }

        /*  /// <summary>
          /// Кнопка для сохранения изменений в строке
          /// </summary>
          /// <param name="sender"></param>
          /// <param name="e"></param>
          private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
          {
              // Игнорируем клик, если в этой ячейке не кнопка
              if (e.RowIndex < 0 || e.ColumnIndex !=
                  dataGridView1.Columns["ColumnOk"].Index) return;
              this.BindingContext[okpoDataSet, "OKPO"].EndCurrentEdit();
              sqlOKPOAdapter.Update(okpoDataSet, "OKPO");
          }    */
    }
}
