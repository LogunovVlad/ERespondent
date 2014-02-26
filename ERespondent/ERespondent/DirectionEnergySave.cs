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
    public partial class DirectionEnergySave : Form
    {
        public DirectionEnergySave()
        {
            InitializeComponent();
            MainForm formMain = this.Owner as MainForm;
        }

        private void DirectionEnergySave_Load(object sender, EventArgs e)
        {
            sqlDirectionConnection.ConnectionString = ConfigurationManager.ConnectionStrings["ERConn"].ConnectionString;
            daDirectSave.Fill(dsDirectSave);
        }
    }
}
