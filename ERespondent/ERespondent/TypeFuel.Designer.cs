using ERespondent.DataSet;
namespace ERespondent
{
    partial class TypeFuel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.codeRecordDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.codeTypeFuelDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.typeFuelDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dsTypeFuel = new ERespondent.DataSet.dsTypeFuel();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.daTypeFuel = new System.Data.SqlClient.SqlDataAdapter();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTypeFuel)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(2, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(777, 448);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.codeRecordDataGridViewTextBoxColumn,
            this.codeTypeFuelDataGridViewTextBoxColumn,
            this.typeFuelDataGridViewTextBoxColumn});
            this.dataGridView1.DataMember = "TypeFuelEnergy";
            this.dataGridView1.DataSource = this.dsTypeFuel;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(771, 442);
            this.dataGridView1.TabIndex = 0;
            // 
            // codeRecordDataGridViewTextBoxColumn
            // 
            this.codeRecordDataGridViewTextBoxColumn.DataPropertyName = "CodeRecord";
            this.codeRecordDataGridViewTextBoxColumn.HeaderText = "CodeRecord";
            this.codeRecordDataGridViewTextBoxColumn.Name = "codeRecordDataGridViewTextBoxColumn";
            this.codeRecordDataGridViewTextBoxColumn.ReadOnly = true;
            this.codeRecordDataGridViewTextBoxColumn.Visible = false;
            // 
            // codeTypeFuelDataGridViewTextBoxColumn
            // 
            this.codeTypeFuelDataGridViewTextBoxColumn.DataPropertyName = "CodeTypeFuel";
            this.codeTypeFuelDataGridViewTextBoxColumn.HeaderText = "Код вида топлива и энергии";
            this.codeTypeFuelDataGridViewTextBoxColumn.Name = "codeTypeFuelDataGridViewTextBoxColumn";
            // 
            // typeFuelDataGridViewTextBoxColumn
            // 
            this.typeFuelDataGridViewTextBoxColumn.DataPropertyName = "TypeFuel";
            this.typeFuelDataGridViewTextBoxColumn.HeaderText = "Вид топлива и энергии";
            this.typeFuelDataGridViewTextBoxColumn.Name = "typeFuelDataGridViewTextBoxColumn";
            // 
            // dsTypeFuel
            // 
            this.dsTypeFuel.DataSetName = "dsTypeFuel";
            this.dsTypeFuel.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "SELECT        CodeRecord, CodeTypeFuel, TypeFuel\r\nFROM            TypeFuelEnergy";
            this.sqlSelectCommand1.Connection = this.sqlConnection1;
            // 
            // sqlConnection1
            // 
            this.sqlConnection1.ConnectionString = "Data Source=VLADPC\\MYSQL05;Initial Catalog=E_Respondent;Integrated Security=True";
            this.sqlConnection1.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlInsertCommand1
            // 
            this.sqlInsertCommand1.CommandText = "INSERT INTO [TypeFuelEnergy] ([CodeTypeFuel], [TypeFuel]) VALUES (@CodeTypeFuel, " +
    "@TypeFuel)";
            this.sqlInsertCommand1.Connection = this.sqlConnection1;
            this.sqlInsertCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@CodeTypeFuel", System.Data.SqlDbType.Int, 0, "CodeTypeFuel"),
            new System.Data.SqlClient.SqlParameter("@TypeFuel", System.Data.SqlDbType.Text, 0, "TypeFuel")});
            // 
            // daTypeFuel
            // 
            this.daTypeFuel.InsertCommand = this.sqlInsertCommand1;
            this.daTypeFuel.SelectCommand = this.sqlSelectCommand1;
            this.daTypeFuel.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "TypeFuelEnergy", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("CodeRecord", "CodeRecord"),
                        new System.Data.Common.DataColumnMapping("CodeTypeFuel", "CodeTypeFuel"),
                        new System.Data.Common.DataColumnMapping("TypeFuel", "TypeFuel")})});
            // 
            // TypeFuel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(784, 462);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "TypeFuel";
            this.Text = "TypeFuel";
            this.Load += new System.EventHandler(this.TypeFuel_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTypeFuel)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private dsTypeFuel dsTypeFuel;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection sqlConnection1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlDataAdapter daTypeFuel;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeRecordDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeTypeFuelDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn typeFuelDataGridViewTextBoxColumn;
    }
}