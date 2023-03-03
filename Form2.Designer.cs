
namespace FemsaTools
{
    partial class Form2
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.txtRE = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button12 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbDivisao = new System.Windows.Forms.ComboBox();
            this.button5 = new System.Windows.Forms.Button();
            this.txtREDivisao = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.lstData = new System.Windows.Forms.ListView();
            this.Data = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Persid = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Tipo = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Persno = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Nome = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.CardNO = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Local = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Mensagem = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.pgBar = new System.Windows.Forms.ProgressBar();
            this.rdbAll = new System.Windows.Forms.RadioButton();
            this.button4 = new System.Windows.Forms.Button();
            this.txtREInfo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.rdbAuthorized = new System.Windows.Forms.RadioButton();
            this.rdbNotYet = new System.Windows.Forms.RadioButton();
            this.rdbAccess = new System.Windows.Forms.RadioButton();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.button11 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.button13 = new System.Windows.Forms.Button();
            this.button14 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(776, 426);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.button6);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(768, 400);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Rotinas WFM";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(6, 88);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(121, 35);
            this.button3.TabIndex = 11;
            this.button3.Text = "IMPORTACÃO PRÓPRIOS - LISTA";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExecute);
            this.groupBox1.Controls.Add(this.txtRE);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(142, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(253, 76);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "WFM Individual";
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(157, 17);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(75, 23);
            this.btnExecute.TabIndex = 2;
            this.btnExecute.Text = "Executar";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // txtRE
            // 
            this.txtRE.Location = new System.Drawing.Point(51, 19);
            this.txtRE.Name = "txtRE";
            this.txtRE.Size = new System.Drawing.Size(100, 20);
            this.txtRE.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "RE;";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 47);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 35);
            this.button1.TabIndex = 9;
            this.button1.Text = "EXCECÕES";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(6, 6);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(121, 35);
            this.button6.TabIndex = 8;
            this.button6.Text = "BLOQUEIOS";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button14);
            this.tabPage2.Controls.Add(this.button13);
            this.tabPage2.Controls.Add(this.button12);
            this.tabPage2.Controls.Add(this.button8);
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(768, 400);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Administração";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(260, 6);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(121, 35);
            this.button12.TabIndex = 12;
            this.button12.Text = "Excluir sem Acesso";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(133, 6);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(121, 35);
            this.button8.TabIndex = 11;
            this.button8.Text = "Importar Empresas";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.cmbDivisao);
            this.groupBox3.Controls.Add(this.button5);
            this.groupBox3.Controls.Add(this.txtREDivisao);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(6, 47);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(266, 100);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Alterar Divisão";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 22);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Divisão:";
            // 
            // cmbDivisao
            // 
            this.cmbDivisao.FormattingEnabled = true;
            this.cmbDivisao.Location = new System.Drawing.Point(79, 19);
            this.cmbDivisao.Name = "cmbDivisao";
            this.cmbDivisao.Size = new System.Drawing.Size(181, 21);
            this.cmbDivisao.TabIndex = 6;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(185, 54);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 5;
            this.button5.Text = "Executar";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // txtREDivisao
            // 
            this.txtREDivisao.Location = new System.Drawing.Point(79, 56);
            this.txtREDivisao.Name = "txtREDivisao";
            this.txtREDivisao.Size = new System.Drawing.Size(100, 20);
            this.txtREDivisao.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 59);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "RE;";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(6, 6);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(121, 35);
            this.button2.TabIndex = 9;
            this.button2.Text = "Adicionar Autorizações";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.lstData);
            this.tabPage3.Controls.Add(this.groupBox2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(768, 400);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Troubleshooting";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // lstData
            // 
            this.lstData.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Data,
            this.Persid,
            this.Tipo,
            this.Persno,
            this.Nome,
            this.CardNO,
            this.Local,
            this.Mensagem});
            this.lstData.HideSelection = false;
            this.lstData.Location = new System.Drawing.Point(12, 131);
            this.lstData.Name = "lstData";
            this.lstData.Size = new System.Drawing.Size(741, 253);
            this.lstData.TabIndex = 11;
            this.lstData.UseCompatibleStateImageBehavior = false;
            // 
            // Data
            // 
            this.Data.Text = "Data";
            // 
            // Persid
            // 
            this.Persid.Text = "Persid";
            // 
            // Tipo
            // 
            this.Tipo.DisplayIndex = 7;
            this.Tipo.Text = "Tipo";
            // 
            // Persno
            // 
            this.Persno.DisplayIndex = 2;
            this.Persno.Text = "Persno";
            // 
            // Nome
            // 
            this.Nome.DisplayIndex = 3;
            this.Nome.Text = "Nome";
            this.Nome.Width = 200;
            // 
            // CardNO
            // 
            this.CardNO.DisplayIndex = 4;
            this.CardNO.Text = "CardNO";
            // 
            // Local
            // 
            this.Local.DisplayIndex = 5;
            this.Local.Text = "Local";
            this.Local.Width = 150;
            // 
            // Mensagem
            // 
            this.Mensagem.DisplayIndex = 6;
            this.Mensagem.Text = "Mensagem";
            this.Mensagem.Width = 150;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.pgBar);
            this.groupBox2.Controls.Add(this.rdbAll);
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Controls.Add(this.txtREInfo);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.rdbAuthorized);
            this.groupBox2.Controls.Add(this.rdbNotYet);
            this.groupBox2.Controls.Add(this.rdbAccess);
            this.groupBox2.Location = new System.Drawing.Point(12, 14);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(524, 100);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Verificar INFO";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(257, 42);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(107, 28);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Exportar";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Visible = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // pgBar
            // 
            this.pgBar.Location = new System.Drawing.Point(6, 73);
            this.pgBar.Name = "pgBar";
            this.pgBar.Size = new System.Drawing.Size(512, 23);
            this.pgBar.Step = 1;
            this.pgBar.TabIndex = 10;
            this.pgBar.Visible = false;
            // 
            // rdbAll
            // 
            this.rdbAll.AutoSize = true;
            this.rdbAll.Checked = true;
            this.rdbAll.Location = new System.Drawing.Point(248, 19);
            this.rdbAll.Name = "rdbAll";
            this.rdbAll.Size = new System.Drawing.Size(55, 17);
            this.rdbAll.TabIndex = 5;
            this.rdbAll.TabStop = true;
            this.rdbAll.Text = "Todos";
            this.rdbAll.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(144, 42);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(107, 28);
            this.button4.TabIndex = 9;
            this.button4.Text = "Executar";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtREInfo
            // 
            this.txtREInfo.Location = new System.Drawing.Point(38, 47);
            this.txtREInfo.Name = "txtREInfo";
            this.txtREInfo.Size = new System.Drawing.Size(100, 20);
            this.txtREInfo.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "RE;";
            // 
            // rdbAuthorized
            // 
            this.rdbAuthorized.AutoSize = true;
            this.rdbAuthorized.Location = new System.Drawing.Point(144, 19);
            this.rdbAuthorized.Name = "rdbAuthorized";
            this.rdbAuthorized.Size = new System.Drawing.Size(98, 17);
            this.rdbAuthorized.TabIndex = 2;
            this.rdbAuthorized.Text = "Não Autorizado";
            this.rdbAuthorized.UseVisualStyleBackColor = true;
            // 
            // rdbNotYet
            // 
            this.rdbNotYet.AutoSize = true;
            this.rdbNotYet.Location = new System.Drawing.Point(72, 19);
            this.rdbNotYet.Name = "rdbNotYet";
            this.rdbNotYet.Size = new System.Drawing.Size(66, 17);
            this.rdbNotYet.TabIndex = 1;
            this.rdbNotYet.Text = "Expirado";
            this.rdbNotYet.UseVisualStyleBackColor = true;
            // 
            // rdbAccess
            // 
            this.rdbAccess.AutoSize = true;
            this.rdbAccess.Location = new System.Drawing.Point(6, 19);
            this.rdbAccess.Name = "rdbAccess";
            this.rdbAccess.Size = new System.Drawing.Size(60, 17);
            this.rdbAccess.TabIndex = 0;
            this.rdbAccess.Text = "Acesso";
            this.rdbAccess.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.button11);
            this.tabPage4.Controls.Add(this.button10);
            this.tabPage4.Controls.Add(this.button9);
            this.tabPage4.Controls.Add(this.button7);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(768, 400);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "SG3";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(420, 14);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(115, 35);
            this.button11.TabIndex = 13;
            this.button11.Text = "Excluir SG3";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(283, 14);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(121, 35);
            this.button10.TabIndex = 12;
            this.button10.Text = "Check SG3 Last Card";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(147, 14);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(121, 35);
            this.button9.TabIndex = 11;
            this.button9.Text = "Check WFM x SG3";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(13, 14);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(121, 35);
            this.button7.TabIndex = 10;
            this.button7.Text = "Importar SG3";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(387, 6);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(121, 35);
            this.button13.TabIndex = 13;
            this.button13.Text = "Ajustar CPF";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(514, 6);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(121, 35);
            this.button14.TabIndex = 14;
            this.button14.Text = "BIS Sem Detector";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Administração BIS";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.TextBox txtRE;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdbAll;
        private System.Windows.Forms.TextBox txtREInfo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton rdbAuthorized;
        private System.Windows.Forms.RadioButton rdbNotYet;
        private System.Windows.Forms.RadioButton rdbAccess;
        private System.Windows.Forms.ListView lstData;
        private System.Windows.Forms.ColumnHeader Data;
        private System.Windows.Forms.ColumnHeader Persid;
        private System.Windows.Forms.ColumnHeader Persno;
        private System.Windows.Forms.ColumnHeader Nome;
        private System.Windows.Forms.ColumnHeader CardNO;
        private System.Windows.Forms.ColumnHeader Local;
        private System.Windows.Forms.ColumnHeader Mensagem;
        private System.Windows.Forms.ProgressBar pgBar;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ColumnHeader Tipo;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbDivisao;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TextBox txtREDivisao;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.Button button14;
    }
}