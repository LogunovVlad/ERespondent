using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERespondent.CheckData
{
    public partial class CheckProtocol : Form
    {
        List<string> list;
        public CheckProtocol(List<String> listError)
        {
            InitializeComponent();
            list = listError;
        }

        private void CheckProtocol_Load(object sender, EventArgs e)
        {
            richTextBox1.ReadOnly = true;
            richTextBox1.Text += "Список ошибок:\n";                    
            if (list.Count!=0)
            {
                foreach (string itemError in list)
                {
                    richTextBox1.Text +=itemError + "\n";
                }
            }
            else
            {                
                richTextBox1.Text = "Ошибки не найдены!";
            }
        }
    }
}
