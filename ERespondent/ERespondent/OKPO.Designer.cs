using ERespondent.DataSet;
namespace ERespondent
{
    partial class OKPO
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OKPO));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btrDel = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this._connectionOKPO = new System.Data.SqlClient.SqlConnection();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
            this._daOKPO = new System.Data.SqlClient.SqlDataAdapter();
            this._dsOKPO = new ERespondent.DataSet._dsOKPO();
            this.codeOKPODataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nameOrganizationDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.codeRecordDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._dsOKPO)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.MinimumSize = new System.Drawing.Size(490, 182);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(816, 375);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.68421F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.89474F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 43.09211F));
            this.tableLayoutPanel2.Controls.Add(this.btnEdit, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.btrDel, 2, 0);
            this.tableLayoutPanel2.Controls.Add(this.btnAdd, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.MinimumSize = new System.Drawing.Size(342, 31);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(342, 44);
            this.tableLayoutPanel2.TabIndex = 2;
            // 
            // btnEdit
            // 
            this.btnEdit.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit.Image")));
            this.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEdit.Location = new System.Drawing.Point(84, 3);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 29);
            this.btnEdit.TabIndex = 1;
            this.btnEdit.Text = "Редактировать";
            this.btnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btrDel
            // 
            this.btrDel.Image = ((System.Drawing.Image)(resources.GetObject("btrDel.Image")));
            this.btrDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btrDel.Location = new System.Drawing.Point(196, 3);
            this.btrDel.Name = "btrDel";
            this.btrDel.Size = new System.Drawing.Size(88, 29);
            this.btrDel.TabIndex = 2;
            this.btrDel.Text = "Удалить";
            this.btrDel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btrDel.UseVisualStyleBackColor = true;
            this.btrDel.Click += new System.EventHandler(this.btrDel_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAdd.Location = new System.Drawing.Point(3, 3);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 29);
            this.btnAdd.TabIndex = 0;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.codeOKPODataGridViewTextBoxColumn,
            this.nameOrganizationDataGridViewTextBoxColumn,
            this.codeRecordDataGridViewTextBoxColumn});
            this.dataGridView1.DataMember = "OKPO";
            this.dataGridView1.DataSource = this._dsOKPO;
            this.dataGridView1.Location = new System.Drawing.Point(3, 53);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(810, 319);
            this.dataGridView1.TabIndex = 1;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(729, 381);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(84, 26);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Закрыть";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.CloseOkpo_Click);
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "SELECT        CodeOKPO, NameOrganization, CodeRecord\r\nFROM            OKPO";
            this.sqlSelectCommand1.Connection = this._connectionOKPO;
            // 
            // _connectionOKPO
            // 
            this._connectionOKPO.ConnectionString = "Data Source=VLADPC\\MYSQL05;Initial Catalog=E_Respondent;Integrated Security=True";
            this._connectionOKPO.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlInsertCommand1
            // 
            this.sqlInsertCommand1.CommandText = "INSERT INTO [OKPO] ([CodeOKPO], [NameOrganization]) VALUES (@CodeOKPO, @NameOrgan" +
    "ization);\r\nSELECT CodeOKPO, NameOrganization, CodeRecord FROM OKPO WHERE (CodeRe" +
    "cord = SCOPE_IDENTITY())";
            this.sqlInsertCommand1.Connection = this._connectionOKPO;
            this.sqlInsertCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@CodeOKPO", System.Data.SqlDbType.Float, 0, "CodeOKPO"),
            new System.Data.SqlClient.SqlParameter("@NameOrganization", System.Data.SqlDbType.Text, 0, "NameOrganization")});
            // 
            // sqlUpdateCommand1
            // 
            this.sqlUpdateCommand1.CommandText = resources.GetString("sqlUpdateCommand1.CommandText");
            this.sqlUpdateCommand1.Connection = this._connectionOKPO;
            this.sqlUpdateCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@CodeOKPO", System.Data.SqlDbType.Float, 0, "CodeOKPO"),
            new System.Data.SqlClient.SqlParameter("@NameOrganization", System.Data.SqlDbType.Text, 0, "NameOrganization"),
            new System.Data.SqlClient.SqlParameter("@IsNull_CodeOKPO", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "CodeOKPO", System.Data.DataRowVersion.Original, true, null, "", "", ""),
            new System.Data.SqlClient.SqlParameter("@Original_CodeOKPO", System.Data.SqlDbType.Float, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CodeOKPO", System.Data.DataRowVersion.Original, null),
            new System.Data.SqlClient.SqlParameter("@Original_CodeRecord", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CodeRecord", System.Data.DataRowVersion.Original, null),
            new System.Data.SqlClient.SqlParameter("@CodeRecord", System.Data.SqlDbType.Int, 4, "CodeRecord")});
            // 
            // sqlDeleteCommand1
            // 
            this.sqlDeleteCommand1.CommandText = "DELETE FROM [OKPO] WHERE (((@IsNull_CodeOKPO = 1 AND [CodeOKPO] IS NULL) OR ([Cod" +
    "eOKPO] = @Original_CodeOKPO)) AND ([CodeRecord] = @Original_CodeRecord))";
            this.sqlDeleteCommand1.Connection = this._connectionOKPO;
            this.sqlDeleteCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@IsNull_CodeOKPO", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, ((byte)(0)), ((byte)(0)), "CodeOKPO", System.Data.DataRowVersion.Original, true, null, "", "", ""),
            new System.Data.SqlClient.SqlParameter("@Original_CodeOKPO", System.Data.SqlDbType.Float, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CodeOKPO", System.Data.DataRowVersion.Original, null),
            new System.Data.SqlClient.SqlParameter("@Original_CodeRecord", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, false, ((byte)(0)), ((byte)(0)), "CodeRecord", System.Data.DataRowVersion.Original, null)});
            // 
            // _daOKPO
            // 
            this._daOKPO.DeleteCommand = this.sqlDeleteCommand1;
            this._daOKPO.InsertCommand = this.sqlInsertCommand1;
            this._daOKPO.SelectCommand = this.sqlSelectCommand1;
            this._daOKPO.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "OKPO", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("CodeOKPO", "CodeOKPO"),
                        new System.Data.Common.DataColumnMapping("NameOrganization", "NameOrganization"),
                        new System.Data.Common.DataColumnMapping("CodeRecord", "CodeRecord")})});
            this._daOKPO.UpdateCommand = this.sqlUpdateCommand1;
            // 
            // _dsOKPO
            // 
            this._dsOKPO.DataSetName = "_dsOKPO";
            this._dsOKPO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // codeOKPODataGridViewTextBoxColumn
            // 
            this.codeOKPODataGridViewTextBoxColumn.DataPropertyName = "CodeOKPO";
            this.codeOKPODataGridViewTextBoxColumn.HeaderText = "Код ОКПО";
            this.codeOKPODataGridViewTextBoxColumn.Name = "codeOKPODataGridViewTextBoxColumn";
            // 
            // nameOrganizationDataGridViewTextBoxColumn
            // 
            this.nameOrganizationDataGridViewTextBoxColumn.DataPropertyName = "NameOrganization";
            this.nameOrganizationDataGridViewTextBoxColumn.HeaderText = "Наименование организации";
            this.nameOrganizationDataGridViewTextBoxColumn.Name = "nameOrganizationDataGridViewTextBoxColumn";
            // 
            // codeRecordDataGridViewTextBoxColumn
            // 
            this.codeRecordDataGridViewTextBoxColumn.DataPropertyName = "CodeRecord";
            this.codeRecordDataGridViewTextBoxColumn.HeaderText = "CodeRecord";
            this.codeRecordDataGridViewTextBoxColumn.Name = "codeRecordDataGridViewTextBoxColumn";
            this.codeRecordDataGridViewTextBoxColumn.ReadOnly = true;
            this.codeRecordDataGridViewTextBoxColumn.Visible = false;
            // 
            // OKPO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(818, 406);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tableLayoutPanel1);
            this.MinimumSize = new System.Drawing.Size(330, 210);
            this.Name = "OKPO";
            this.Text = "Справочник кодов ОКПО организаций (подразделений)";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.OKPO_FormClosed);
            this.Load += new System.EventHandler(this.OKPO_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._dsOKPO)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btrDel;
        public System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        protected internal System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnCancel;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection _connectionOKPO;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
        private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;
        protected internal System.Data.SqlClient.SqlDataAdapter _daOKPO;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeOKPODataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nameOrganizationDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn codeRecordDataGridViewTextBoxColumn;
        protected internal _dsOKPO _dsOKPO;
    }
}