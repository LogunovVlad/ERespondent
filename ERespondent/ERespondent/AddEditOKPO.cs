using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERespondent
{
    public partial class AddEditOKPO : Form
    {
        public AddEditOKPO()
        {
            InitializeComponent();
        }
        OKPO formOKPO;
        private void AddEditOKPO_Load(object sender, EventArgs e)
        {
            formOKPO = this.Owner as OKPO;
            //проверка: какую функцию давать форме - добавление или редактирование
            if (formOKPO.Tag.Equals("ADD"))
            {
                btnOk.Click += new EventHandler(ClickAdd);
            }
            else
            {
                try
                {
                    txtOkpo.Text = formOKPO.dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    txtName.Text = formOKPO.dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    btnOk.Click += new EventHandler(ClickEdit);
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Не выбрана запись для редактирования!", "Ошибка редактирования!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        /// <summary>
        /// Редактирование записи в DataTable
        /// </summary>
        private void ClickEdit(object sender, EventArgs e)
        {
            try
            {
                //определим индекс изменяемой строки
                int indexRow = formOKPO.dataGridView1.CurrentRow.Index;
                formOKPO._dsOKPO.Tables["OKPO"].Rows[indexRow]["CodeOKPO"] = txtOkpo.Text.ToString();
                formOKPO._dsOKPO.Tables["OKPO"].Rows[indexRow]["NameOrganization"] = txtName.Text.ToString();

                formOKPO._daOKPO.Update(formOKPO._dsOKPO);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Добавление записи в DataSet, путем создания DataRow и вставки этой строки DataTable
        /// </summary>        
        private void ClickAdd(object sender, EventArgs e)
        {
            try
            {
                DataTable tblOKPO;
                tblOKPO = formOKPO._dsOKPO.Tables["OKPO"];
                DataRow newOKPORow;
                newOKPORow = tblOKPO.NewRow();
                newOKPORow["CodeOKPO"] = txtOkpo.Text;
                newOKPORow["NameOrganization"] = txtName.Text;
                tblOKPO.Rows.Add(newOKPORow);

                formOKPO._daOKPO.Update(formOKPO._dsOKPO);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
