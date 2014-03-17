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
        public static Dictionary<string, List<string>> ErrorForAllSection;
        public CheckProtocol()
        {
            InitializeComponent();
            ErrorForAllSection = new Dictionary<string, List<string>>();
        }

        private void CheckProtocol_Load(object sender, EventArgs e)
        {
            richTextBox1.ReadOnly = true;
            richTextBox1.Text += "Список ошибок:\n";
            for (int i = 0; i < ErrorForAllSection.Count; i++)
            {
                richTextBox1.Text += ErrorForAllSection.ElementAt(i).Key;
                if (ErrorForAllSection.ElementAt(i).Value.Count > 0)
                {
                    foreach (string itemError in ErrorForAllSection.ElementAt(i).Value)
                    {
                        richTextBox1.Text += itemError + "\n";
                    }
                }
                else
                {
                    richTextBox1.Text += "Ошибок не найдено!";
                }
            }
        }
    }
}
