namespace SpecificationPack
{
    partial class MainForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.specListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.addSpecBtn = new System.Windows.Forms.Button();
            this.deleteSpecBtn = new System.Windows.Forms.Button();
            this.formBtn = new System.Windows.Forms.Button();
            this.clearSpecBtn = new System.Windows.Forms.Button();
            this.groupCheckBox = new System.Windows.Forms.CheckBox();
            this.unionCheckBox = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // specListBox
            // 
            this.specListBox.AllowDrop = true;
            this.specListBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.specListBox.FormattingEnabled = true;
            this.specListBox.Location = new System.Drawing.Point(12, 28);
            this.specListBox.Name = "specListBox";
            this.specListBox.Size = new System.Drawing.Size(499, 147);
            this.specListBox.TabIndex = 0;
            this.specListBox.DragDrop += new System.Windows.Forms.DragEventHandler(this.specListBox_DragDrop);
            this.specListBox.DragEnter += new System.Windows.Forms.DragEventHandler(this.specListBox_DragEnter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Спецификации";
            // 
            // addSpecBtn
            // 
            this.addSpecBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.addSpecBtn.Location = new System.Drawing.Point(412, 177);
            this.addSpecBtn.Name = "addSpecBtn";
            this.addSpecBtn.Size = new System.Drawing.Size(99, 36);
            this.addSpecBtn.TabIndex = 2;
            this.addSpecBtn.Text = "Добавить спецификацию";
            this.addSpecBtn.UseVisualStyleBackColor = true;
            this.addSpecBtn.Click += new System.EventHandler(this.addSpecBtn_Click);
            // 
            // deleteSpecBtn
            // 
            this.deleteSpecBtn.Location = new System.Drawing.Point(12, 177);
            this.deleteSpecBtn.Name = "deleteSpecBtn";
            this.deleteSpecBtn.Size = new System.Drawing.Size(99, 36);
            this.deleteSpecBtn.TabIndex = 2;
            this.deleteSpecBtn.Text = "Удалить";
            this.deleteSpecBtn.UseVisualStyleBackColor = true;
            this.deleteSpecBtn.Click += new System.EventHandler(this.deleteSpecBtn_Click);
            // 
            // formBtn
            // 
            this.formBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.formBtn.Location = new System.Drawing.Point(393, 224);
            this.formBtn.Name = "formBtn";
            this.formBtn.Size = new System.Drawing.Size(118, 34);
            this.formBtn.TabIndex = 3;
            this.formBtn.Text = "Сформировать";
            this.formBtn.UseVisualStyleBackColor = true;
            this.formBtn.Click += new System.EventHandler(this.formBtn_Click);
            // 
            // clearSpecBtn
            // 
            this.clearSpecBtn.Location = new System.Drawing.Point(117, 177);
            this.clearSpecBtn.Name = "clearSpecBtn";
            this.clearSpecBtn.Size = new System.Drawing.Size(99, 36);
            this.clearSpecBtn.TabIndex = 2;
            this.clearSpecBtn.Text = "Очистить";
            this.clearSpecBtn.UseVisualStyleBackColor = true;
            this.clearSpecBtn.Click += new System.EventHandler(this.clearSpecBtn_Click);
            // 
            // groupCheckBox
            // 
            this.groupCheckBox.AutoSize = true;
            this.groupCheckBox.Location = new System.Drawing.Point(12, 233);
            this.groupCheckBox.Name = "groupCheckBox";
            this.groupCheckBox.Size = new System.Drawing.Size(118, 17);
            this.groupCheckBox.TabIndex = 4;
            this.groupCheckBox.Text = "Добавлять группу";
            this.groupCheckBox.UseVisualStyleBackColor = true;
            // 
            // unionCheckBox
            // 
            this.unionCheckBox.AutoSize = true;
            this.unionCheckBox.Location = new System.Drawing.Point(12, 257);
            this.unionCheckBox.Name = "unionCheckBox";
            this.unionCheckBox.Size = new System.Drawing.Size(115, 17);
            this.unionCheckBox.TabIndex = 5;
            this.unionCheckBox.Text = "Без объединения";
            this.unionCheckBox.UseVisualStyleBackColor = true;
            this.unionCheckBox.CheckedChanged += new System.EventHandler(this.unionCheckBox_CheckedChanged);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Yellow;
            this.panel1.Location = new System.Drawing.Point(16, 303);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(20, 18);
            this.panel1.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(42, 306);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(289, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Совпадают коды, но не совпадают единицы измерения";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(42, 333);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(260, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Совпадают наименования, но не совпадают коды";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Magenta;
            this.panel2.Location = new System.Drawing.Point(16, 330);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(20, 18);
            this.panel2.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(42, 361);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(471, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Отсутствует код, но совпадают наименования и данного наименования нет в исключени" +
    "ях";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Red;
            this.panel3.Location = new System.Drawing.Point(16, 358);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(20, 18);
            this.panel3.TabIndex = 10;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(523, 386);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.unionCheckBox);
            this.Controls.Add(this.groupCheckBox);
            this.Controls.Add(this.formBtn);
            this.Controls.Add(this.clearSpecBtn);
            this.Controls.Add(this.deleteSpecBtn);
            this.Controls.Add(this.addSpecBtn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.specListBox);
            this.Name = "MainForm";
            this.Text = "Расчёт количества упаковок";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox specListBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button addSpecBtn;
        private System.Windows.Forms.Button deleteSpecBtn;
        private System.Windows.Forms.Button formBtn;
        private System.Windows.Forms.Button clearSpecBtn;
        private System.Windows.Forms.CheckBox groupCheckBox;
        private System.Windows.Forms.CheckBox unionCheckBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel3;
    }
}

