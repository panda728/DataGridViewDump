namespace DataGridViewDump
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet11 = new DataGridViewDump.DataSet1();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.col1DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col2DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col3DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet11)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // bindingSource1
            // 
            this.bindingSource1.AllowNew = false;
            this.bindingSource1.DataMember = "Table1";
            this.bindingSource1.DataSource = this.dataSet11;
            // 
            // dataSet11
            // 
            this.dataSet11.DataSetName = "DataSet1";
            this.dataSet11.Namespace = "http://tempuri.org/DataSet1.xsd";
            this.dataSet11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.col1DataGridViewTextBoxColumn,
            this.col2DataGridViewTextBoxColumn,
            this.col3DataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.bindingSource1;
            this.dataGridView1.Location = new System.Drawing.Point(12, 68);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(776, 286);
            this.dataGridView1.TabIndex = 0;
            // 
            // col1DataGridViewTextBoxColumn
            // 
            this.col1DataGridViewTextBoxColumn.DataPropertyName = "Col1";
            this.col1DataGridViewTextBoxColumn.HeaderText = "Col1";
            this.col1DataGridViewTextBoxColumn.Name = "col1DataGridViewTextBoxColumn";
            // 
            // col2DataGridViewTextBoxColumn
            // 
            this.col2DataGridViewTextBoxColumn.DataPropertyName = "Col2";
            this.col2DataGridViewTextBoxColumn.HeaderText = "Col2";
            this.col2DataGridViewTextBoxColumn.Name = "col2DataGridViewTextBoxColumn";
            // 
            // col3DataGridViewTextBoxColumn
            // 
            this.col3DataGridViewTextBoxColumn.DataPropertyName = "Col3";
            this.col3DataGridViewTextBoxColumn.HeaderText = "Col3";
            this.col3DataGridViewTextBoxColumn.Name = "col3DataGridViewTextBoxColumn";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(212, 34);
            this.button1.TabIndex = 1;
            this.button1.Text = "DataGridView to ExcelFile";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet11)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private BindingSource bindingSource1;
        private DataGridView dataGridView1;
        private Button button1;
        private SaveFileDialog saveFileDialog1;
        private DataSet1 dataSet11;
        private DataGridViewTextBoxColumn col1DataGridViewTextBoxColumn;
        private DataGridViewTextBoxColumn col2DataGridViewTextBoxColumn;
        private DataGridViewTextBoxColumn col3DataGridViewTextBoxColumn;
    }
}