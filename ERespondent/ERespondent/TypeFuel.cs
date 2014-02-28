using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace ERespondent
{
    public partial class TypeFuel : Form
    {
        public TypeFuel()
        {
            InitializeComponent();
        }

        private void TypeFuel_Load(object sender, EventArgs e)
        {
            sqlConnection1.ConnectionString = ConfigurationManager.ConnectionStrings["ERConn"].ConnectionString;
            MainForm mainForm=this.Owner as MainForm;
            daTypeFuel.Fill(dsTypeFuel);
        }
    }
}
