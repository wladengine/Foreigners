namespace Priem
{
    partial class OrderNumbersList
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OrderNumbers));
            this.dgvViews = new System.Windows.Forms.DataGridView();
            this.gbOrders = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tbOrderNum = new System.Windows.Forms.TextBox();
            this.dtOrderDate = new System.Windows.Forms.DateTimePicker();
            this.gbOrdersFor = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.tbOrderNumFor = new System.Windows.Forms.TextBox();
            this.dtOrderDateFor = new System.Windows.Forms.DateTimePicker();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnChange = new System.Windows.Forms.Button();
            this.chbIsSecond = new System.Windows.Forms.CheckBox();
            this.chbIsReduced = new System.Windows.Forms.CheckBox();
            this.chbIsParallel = new System.Windows.Forms.CheckBox();
            this.cbStudyForm = new System.Windows.Forms.ComboBox();
            this.cbStudyBasis = new System.Windows.Forms.ComboBox();
            this.cbLicenseProgram = new System.Windows.Forms.ComboBox();
            this.cbFaculty = new System.Windows.Forms.ComboBox();
            this.chbIsListener = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvViews)).BeginInit();
            this.gbOrders.SuspendLayout();
            this.gbOrdersFor.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvViews
            // 
            this.dgvViews.AllowUserToAddRows = false;
            this.dgvViews.AllowUserToDeleteRows = false;
            this.dgvViews.AllowUserToResizeRows = false;
            this.dgvViews.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvViews.Location = new System.Drawing.Point(6, 257);
            this.dgvViews.MultiSelect = false;
            this.dgvViews.Name = "dgvViews";
            this.dgvViews.ReadOnly = true;
            this.dgvViews.RowHeadersVisible = false;
            this.dgvViews.Size = new System.Drawing.Size(152, 225);
            this.dgvViews.TabIndex = 83;
            this.dgvViews.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvViews_RowEnter);
            // 
            // gbOrders
            // 
            this.gbOrders.Controls.Add(this.label5);
            this.gbOrders.Controls.Add(this.label1);
            this.gbOrders.Controls.Add(this.tbOrderNum);
            this.gbOrders.Controls.Add(this.dtOrderDate);
            this.gbOrders.Location = new System.Drawing.Point(183, 257);
            this.gbOrders.Name = "gbOrders";
            this.gbOrders.Size = new System.Drawing.Size(162, 109);
            this.gbOrders.TabIndex = 84;
            this.gbOrders.TabStop = false;
            this.gbOrders.Text = "Приказ для граждан РФ";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 67);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 13);
            this.label5.TabIndex = 87;
            this.label5.Text = "Номер";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 86;
            this.label1.Text = "Дата";
            // 
            // tbOrderNum
            // 
            this.tbOrderNum.Location = new System.Drawing.Point(9, 83);
            this.tbOrderNum.Name = "tbOrderNum";
            this.tbOrderNum.Size = new System.Drawing.Size(131, 20);
            this.tbOrderNum.TabIndex = 85;
            // 
            // dtOrderDate
            // 
            this.dtOrderDate.Location = new System.Drawing.Point(9, 36);
            this.dtOrderDate.Name = "dtOrderDate";
            this.dtOrderDate.Size = new System.Drawing.Size(131, 20);
            this.dtOrderDate.TabIndex = 0;
            // 
            // gbOrdersFor
            // 
            this.gbOrdersFor.Controls.Add(this.label6);
            this.gbOrdersFor.Controls.Add(this.label4);
            this.gbOrdersFor.Controls.Add(this.tbOrderNumFor);
            this.gbOrdersFor.Controls.Add(this.dtOrderDateFor);
            this.gbOrdersFor.Location = new System.Drawing.Point(183, 373);
            this.gbOrdersFor.Name = "gbOrdersFor";
            this.gbOrdersFor.Size = new System.Drawing.Size(162, 109);
            this.gbOrdersFor.TabIndex = 85;
            this.gbOrdersFor.TabStop = false;
            this.gbOrdersFor.Text = "Приказ для иностранцев";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 25);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 13);
            this.label6.TabIndex = 87;
            this.label6.Text = "Дата";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 13);
            this.label4.TabIndex = 86;
            this.label4.Text = "Номер";
            // 
            // tbOrderNumFor
            // 
            this.tbOrderNumFor.Location = new System.Drawing.Point(9, 83);
            this.tbOrderNumFor.Name = "tbOrderNumFor";
            this.tbOrderNumFor.Size = new System.Drawing.Size(131, 20);
            this.tbOrderNumFor.TabIndex = 85;
            // 
            // dtOrderDateFor
            // 
            this.dtOrderDateFor.Location = new System.Drawing.Point(9, 41);
            this.dtOrderDateFor.Name = "dtOrderDateFor";
            this.dtOrderDateFor.Size = new System.Drawing.Size(131, 20);
            this.dtOrderDateFor.TabIndex = 0;
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(280, 494);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 86;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(6, 494);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 87;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnChange
            // 
            this.btnChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChange.Location = new System.Drawing.Point(193, 494);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(75, 23);
            this.btnChange.TabIndex = 88;
            this.btnChange.Text = "Изменить";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // chbIsSecond
            // 
            this.chbIsSecond.AutoSize = true;
            this.chbIsSecond.Location = new System.Drawing.Point(217, 105);
            this.chbIsSecond.Name = "chbIsSecond";
            this.chbIsSecond.Size = new System.Drawing.Size(137, 17);
            this.chbIsSecond.TabIndex = 162;
            this.chbIsSecond.Text = "для лиц, имеющих ВО";
            this.chbIsSecond.UseVisualStyleBackColor = true;
            this.chbIsSecond.CheckedChanged += new System.EventHandler(this.chbIsSecond_CheckedChanged);
            // 
            // chbIsReduced
            // 
            this.chbIsReduced.AutoSize = true;
            this.chbIsReduced.Location = new System.Drawing.Point(12, 105);
            this.chbIsReduced.Name = "chbIsReduced";
            this.chbIsReduced.Size = new System.Drawing.Size(95, 17);
            this.chbIsReduced.TabIndex = 161;
            this.chbIsReduced.Text = "сокращенная";
            this.chbIsReduced.UseVisualStyleBackColor = true;
            this.chbIsReduced.CheckedChanged += new System.EventHandler(this.chbIsReduced_CheckedChanged);
            // 
            // chbIsParallel
            // 
            this.chbIsParallel.AutoSize = true;
            this.chbIsParallel.Location = new System.Drawing.Point(113, 105);
            this.chbIsParallel.Name = "chbIsParallel";
            this.chbIsParallel.Size = new System.Drawing.Size(98, 17);
            this.chbIsParallel.TabIndex = 160;
            this.chbIsParallel.Text = "параллельная";
            this.chbIsParallel.UseVisualStyleBackColor = true;
            this.chbIsParallel.CheckedChanged += new System.EventHandler(this.chbIsParallel_CheckedChanged);
            // 
            // cbStudyForm
            // 
            this.cbStudyForm.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbStudyForm.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbStudyForm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyForm.FormattingEnabled = true;
            this.cbStudyForm.Location = new System.Drawing.Point(12, 147);
            this.cbStudyForm.Name = "cbStudyForm";
            this.cbStudyForm.Size = new System.Drawing.Size(170, 21);
            this.cbStudyForm.TabIndex = 159;
            // 
            // cbStudyBasis
            // 
            this.cbStudyBasis.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbStudyBasis.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbStudyBasis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbStudyBasis.FormattingEnabled = true;
            this.cbStudyBasis.Location = new System.Drawing.Point(12, 68);
            this.cbStudyBasis.Name = "cbStudyBasis";
            this.cbStudyBasis.Size = new System.Drawing.Size(156, 21);
            this.cbStudyBasis.TabIndex = 158;
            // 
            // cbLicenseProgram
            // 
            this.cbLicenseProgram.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbLicenseProgram.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbLicenseProgram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLicenseProgram.FormattingEnabled = true;
            this.cbLicenseProgram.Location = new System.Drawing.Point(12, 192);
            this.cbLicenseProgram.Name = "cbLicenseProgram";
            this.cbLicenseProgram.Size = new System.Drawing.Size(342, 21);
            this.cbLicenseProgram.TabIndex = 157;
            // 
            // cbFaculty
            // 
            this.cbFaculty.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbFaculty.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFaculty.FormattingEnabled = true;
            this.cbFaculty.Location = new System.Drawing.Point(12, 21);
            this.cbFaculty.Name = "cbFaculty";
            this.cbFaculty.Size = new System.Drawing.Size(342, 21);
            this.cbFaculty.TabIndex = 156;
            // 
            // chbIsListener
            // 
            this.chbIsListener.AutoSize = true;
            this.chbIsListener.Location = new System.Drawing.Point(12, 228);
            this.chbIsListener.Name = "chbIsListener";
            this.chbIsListener.Size = new System.Drawing.Size(80, 17);
            this.chbIsListener.TabIndex = 155;
            this.chbIsListener.Text = "слушатели";
            this.chbIsListener.UseVisualStyleBackColor = true;
            this.chbIsListener.CheckedChanged += new System.EventHandler(this.chbIsListener_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 176);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 13);
            this.label2.TabIndex = 154;
            this.label2.Text = "Направление (Направление)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 153;
            this.label3.Text = "Факультет";
            // 
            // lblLanguage
            // 
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(9, 52);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(94, 13);
            this.lblLanguage.TabIndex = 152;
            this.lblLanguage.Text = "Основа обучения";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 131);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(93, 13);
            this.label7.TabIndex = 151;
            this.label7.Text = "Форма обучения";
            // 
            // OrderNumbers
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(366, 529);
            this.Controls.Add(this.chbIsSecond);
            this.Controls.Add(this.chbIsReduced);
            this.Controls.Add(this.chbIsParallel);
            this.Controls.Add(this.cbStudyForm);
            this.Controls.Add(this.cbStudyBasis);
            this.Controls.Add(this.cbLicenseProgram);
            this.Controls.Add(this.cbFaculty);
            this.Controls.Add(this.chbIsListener);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.gbOrdersFor);
            this.Controls.Add(this.gbOrders);
            this.Controls.Add(this.dgvViews);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "OrderNumbers";
            this.Text = "Номера приказов";
            ((System.ComponentModel.ISupportInitialize)(this.dgvViews)).EndInit();
            this.gbOrders.ResumeLayout(false);
            this.gbOrders.PerformLayout();
            this.gbOrdersFor.ResumeLayout(false);
            this.gbOrdersFor.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvViews;
        private System.Windows.Forms.GroupBox gbOrders;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbOrderNum;
        private System.Windows.Forms.DateTimePicker dtOrderDate;
        private System.Windows.Forms.GroupBox gbOrdersFor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbOrderNumFor;
        private System.Windows.Forms.DateTimePicker dtOrderDateFor;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnChange;
        private System.Windows.Forms.CheckBox chbIsSecond;
        private System.Windows.Forms.CheckBox chbIsReduced;
        private System.Windows.Forms.CheckBox chbIsParallel;
        private System.Windows.Forms.ComboBox cbStudyForm;
        private System.Windows.Forms.ComboBox cbStudyBasis;
        private System.Windows.Forms.ComboBox cbLicenseProgram;
        private System.Windows.Forms.ComboBox cbFaculty;
        private System.Windows.Forms.CheckBox chbIsListener;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.Label label7;
    }
}