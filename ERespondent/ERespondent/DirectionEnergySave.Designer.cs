namespace ERespondent
{
    partial class DirectionEnergySave
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
            this.codeDirectionDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.destinationsSaveDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.unitDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dsDirectSave = new ERespondent.dsDirectSave();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDirectionConnection = new System.Data.SqlClient.SqlConnection();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.daDirectSave = new System.Data.SqlClient.SqlDataAdapter();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsDirectSave)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(760, 438);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.codeRecordDataGridViewTextBoxColumn,
            this.codeDirectionDataGridViewTextBoxColumn,
            this.destinationsSaveDataGridViewTextBoxColumn,
            this.unitDataGridViewTextBoxColumn});
            this.dataGridView1.DataMember = "DestinationSave";
            this.dataGridView1.DataSource = this.dsDirectSave;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(754, 432);
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
            // codeDirectionDataGridViewTextBoxColumn
            // 
            this.codeDirectionDataGridViewTextBoxColumn.DataPropertyName = "CodeDirection";
            this.codeDirectionDataGridViewTextBoxColumn.HeaderText = "CodeDirection";
            this.codeDirectionDataGridViewTextBoxColumn.Name = "codeDirectionDataGridViewTextBoxColumn";
            // 
            // destinationsSaveDataGridViewTextBoxColumn
            // 
            this.destinationsSaveDataGridViewTextBoxColumn.DataPropertyName = "DestinationsSave";
            this.destinationsSaveDataGridViewTextBoxColumn.HeaderText = "DestinationsSave";
            this.destinationsSaveDataGridViewTextBoxColumn.Name = "destinationsSaveDataGridViewTextBoxColumn";
            // 
            // unitDataGridViewTextBoxColumn
            // 
            this.unitDataGridViewTextBoxColumn.DataPropertyName = "Unit";
            this.unitDataGridViewTextBoxColumn.HeaderText = "Unit";
            this.unitDataGridViewTextBoxColumn.Name = "unitDataGridViewTextBoxColumn";
            // 
            // dsDirectSave
            // 
            this.dsDirectSave.DataSetName = "dsDirectSave";
            this.dsDirectSave.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "SELECT        CodeRecord, CodeDirection, DestinationsSave, Unit\r\nFROM            " +
    "DestinationSave";
            this.sqlSelectCommand1.Connection = this.sqlDirectionConnection;
            // 
            // sqlDirectionConnection
            // 
            this.sqlDirectionConnection.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlInsertCommand1
            // 
            this.sqlInsertCommand1.CommandText = "INSERT INTO [DestinationSave] ([CodeDirection], [DestinationsSave], [Unit]) VALUE" +
    "S (@CodeDirection, @DestinationsSave, @Unit)";
            this.sqlInsertCommand1.Connection = this.sqlDirectionConnection;
            this.sqlInsertCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@CodeDirection", System.Data.SqlDbType.Int, 0, "CodeDirection"),
            new System.Data.SqlClient.SqlParameter("@DestinationsSave", System.Data.SqlDbType.VarChar, 0, "DestinationsSave"),
            new System.Data.SqlClient.SqlParameter("@Unit", System.Data.SqlDbType.VarChar, 0, "Unit")});
            // 
            // daDirectSave
            // 
            this.daDirectSave.InsertCommand = this.sqlInsertCommand1;
            this.daDirectSave.SelectCommand = this.sqlSelectCommand1;
            this.daDirectSave.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "DestinationSave", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("CodeRecord", "CodeRecord"),
                        new System.Data.Common.DataColumnMapping("CodeDirection", "CodeDirection"),
                        new System.Data.Common.DataColumnMapping("DestinationsSave", "DestinationsSave"),
                        new System.Data.Common.DataColumnMapping("Unit", "Unit")})});
            // 
            // DirectionEnergySave
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 462);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "DirectionEnergySave";
            this.Text = "DirectionEnergySave";
            this.Load += new System.EventHandler(this.DirectionEnergySave_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsDirectSave)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeRecordDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeDirectionDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn destinationsSaveDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn unitDataGridViewTextBoxColumn;
        private dsDirectSave dsDirectSave;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection sqlDirectionConnection;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlDataAdapter daDirectSave;
    }
}