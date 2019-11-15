//----------------------------------------------------------------
// OPCHDA.NET Client Sample Application
// ------------------------------------
// Shows examples of how the OPC server can be browsed and items
// displayed and selected.
//
// Uses the OPCDA.NET Wrapper Assembly to access the OPC HDA Server.
// The items are browsed and the selected item is read in synchronous mode.
//
// Copyright 2003-06 Advosol Inc.   ( www.advosol.com)
// All rights reserved.
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
//----------------------------------------------------------------

using OPC;
using OPCHDA;
using OPCHDA.NET;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ReadRawSyncSample
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {
        public int CountOfDate;
        internal CheckBox ckbBrowseAllInitially;
        private Button btn_export;
        private System.Windows.Forms.Button btnBrowseOPCServers;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnDisconnect;
        private System.Windows.Forms.Button btnRead;
        private Button btnReset;
        private ShowHDABrowseTree BTree = null;
        private System.Windows.Forms.ComboBox cbHostName;
        private System.Windows.Forms.ComboBox cbOPCServers;
        private CheckBox chkBounds;
        private IContainer components;
        private DateTimePicker dtp_startTime;
        private DateTimePicker dtp_stopTime;
        private FlowLayoutPanel flowLayoutPanel1;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private GroupBox groupBox4;
        private Label label2;
        private Label label3;
        private Label label4;
        private Panel mainPanel;
        private SplitContainer mainSplit;
        private NumericUpDown nup_numValues;

        //================================================================
        // OPC Server Access Sample Code
        //------------------------------
        private OpcHDAServer OpcSrv = null;

        private Panel panel1;
        private Panel panel2;
        private Panel panel3;
        private Panel panel4;
        private SplitContainer secondSplit;
        private TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TextBox tbStatus;
        private System.Windows.Forms.TextBox tbValues;
        private NumericUpDown timeInterval;
        private System.Windows.Forms.TreeView tvServer;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Form1());
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
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
            this.btnBrowseOPCServers = new System.Windows.Forms.Button();
            this.cbOPCServers = new System.Windows.Forms.ComboBox();
            this.tvServer = new System.Windows.Forms.TreeView();
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnDisconnect = new System.Windows.Forms.Button();
            this.tbStatus = new System.Windows.Forms.TextBox();
            this.cbHostName = new System.Windows.Forms.ComboBox();
            this.btnRead = new System.Windows.Forms.Button();
            this.tbValues = new System.Windows.Forms.TextBox();
            this.btn_export = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkBounds = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.timeInterval = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.dtp_startTime = new System.Windows.Forms.DateTimePicker();
            this.dtp_stopTime = new System.Windows.Forms.DateTimePicker();
            this.nup_numValues = new System.Windows.Forms.NumericUpDown();
            this.ckbBrowseAllInitially = new System.Windows.Forms.CheckBox();
            this.mainSplit = new System.Windows.Forms.SplitContainer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.secondSplit = new System.Windows.Forms.SplitContainer();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnReset = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.mainPanel = new System.Windows.Forms.Panel();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.timeInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nup_numValues)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainSplit)).BeginInit();
            this.mainSplit.Panel1.SuspendLayout();
            this.mainSplit.Panel2.SuspendLayout();
            this.mainSplit.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.secondSplit)).BeginInit();
            this.secondSplit.Panel1.SuspendLayout();
            this.secondSplit.Panel2.SuspendLayout();
            this.secondSplit.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.panel2.SuspendLayout();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            //
            // btnBrowseOPCServers
            //
            this.btnBrowseOPCServers.Location = new System.Drawing.Point(3, 3);
            this.btnBrowseOPCServers.Name = "btnBrowseOPCServers";
            this.btnBrowseOPCServers.Size = new System.Drawing.Size(102, 28);
            this.btnBrowseOPCServers.TabIndex = 0;
            this.btnBrowseOPCServers.Text = "瀏覽HDA Servers";
            this.btnBrowseOPCServers.Click += new System.EventHandler(this.btnBrowseOPCServers_Click);
            //
            // cbOPCServers
            //
            this.cbOPCServers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cbOPCServers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.cbOPCServers.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbOPCServers.Location = new System.Drawing.Point(0, 0);
            this.cbOPCServers.Name = "cbOPCServers";
            this.cbOPCServers.Size = new System.Drawing.Size(226, 286);
            this.cbOPCServers.Sorted = true;
            this.cbOPCServers.TabIndex = 2;
            //
            // tvServer
            //
            this.tvServer.CheckBoxes = true;
            this.tvServer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvServer.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tvServer.Location = new System.Drawing.Point(0, 0);
            this.tvServer.Name = "tvServer";
            this.tvServer.PathSeparator = ".";
            this.tvServer.Size = new System.Drawing.Size(139, 347);
            this.tvServer.TabIndex = 4;
            //
            // btnConnect
            //
            this.btnConnect.Location = new System.Drawing.Point(111, 3);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(52, 28);
            this.btnConnect.TabIndex = 5;
            this.btnConnect.Text = "連線";
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            //
            // btnDisconnect
            //
            this.btnDisconnect.Enabled = false;
            this.btnDisconnect.Location = new System.Drawing.Point(169, 3);
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.Size = new System.Drawing.Size(52, 28);
            this.btnDisconnect.TabIndex = 6;
            this.btnDisconnect.Text = "離線";
            this.btnDisconnect.Click += new System.EventHandler(this.btnDisconnect_Click);
            //
            // tbStatus
            //
            this.tbStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbStatus.Location = new System.Drawing.Point(0, 0);
            this.tbStatus.Name = "tbStatus";
            this.tbStatus.ReadOnly = true;
            this.tbStatus.Size = new System.Drawing.Size(744, 22);
            this.tbStatus.TabIndex = 7;
            //
            // cbHostName
            //
            this.cbHostName.Location = new System.Drawing.Point(3, 21);
            this.cbHostName.Name = "cbHostName";
            this.cbHostName.Size = new System.Drawing.Size(142, 20);
            this.cbHostName.TabIndex = 21;
            //
            // btnRead
            //
            this.btnRead.Location = new System.Drawing.Point(19, 53);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(128, 27);
            this.btnRead.TabIndex = 23;
            this.btnRead.Text = "讀取所有選擇項目";
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            //
            // tbValues
            //
            this.tableLayoutPanel1.SetColumnSpan(this.tbValues, 2);
            this.tbValues.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbValues.Location = new System.Drawing.Point(3, 191);
            this.tbValues.Multiline = true;
            this.tbValues.Name = "tbValues";
            this.tbValues.ReadOnly = true;
            this.tbValues.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbValues.Size = new System.Drawing.Size(363, 176);
            this.tbValues.TabIndex = 28;
            this.tbValues.WordWrap = false;
            //
            // btn_export
            //
            this.btn_export.Location = new System.Drawing.Point(6, 53);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(157, 27);
            this.btn_export.TabIndex = 30;
            this.btn_export.Text = "匯出 csv";
            this.btn_export.UseVisualStyleBackColor = true;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            //
            // label2
            //
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 12);
            this.label2.TabIndex = 32;
            this.label2.Text = "時間間隔(秒)";
            //
            // label4
            //
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(174, 28);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(11, 12);
            this.label4.TabIndex = 34;
            this.label4.Text = "~";
            //
            // groupBox1
            //
            this.groupBox1.Controls.Add(this.btnRead);
            this.groupBox1.Controls.Add(this.chkBounds);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(3, 103);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(178, 82);
            this.groupBox1.TabIndex = 39;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "項目讀取";
            //
            // chkBounds
            //
            this.chkBounds.Location = new System.Drawing.Point(19, 29);
            this.chkBounds.Name = "chkBounds";
            this.chkBounds.Size = new System.Drawing.Size(78, 18);
            this.chkBounds.TabIndex = 29;
            this.chkBounds.Text = "Bounds";
            //
            // groupBox2
            //
            this.groupBox2.Controls.Add(this.timeInterval);
            this.groupBox2.Controls.Add(this.btn_export);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(187, 103);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(179, 82);
            this.groupBox2.TabIndex = 40;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "匯出";
            //
            // timeInterval
            //
            this.timeInterval.Location = new System.Drawing.Point(83, 25);
            this.timeInterval.Maximum = new decimal(new int[] {
            36400,
            0,
            0,
            0});
            this.timeInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.timeInterval.Name = "timeInterval";
            this.timeInterval.Size = new System.Drawing.Size(80, 22);
            this.timeInterval.TabIndex = 33;
            this.timeInterval.Value = new decimal(new int[] {
            30,
            0,
            0,
            0});
            //
            // label3
            //
            this.label3.Location = new System.Drawing.Point(108, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 10);
            this.label3.TabIndex = 42;
            this.label3.Text = "Number of Values";
            //
            // dtp_startTime
            //
            this.dtp_startTime.CustomFormat = "MM/dd/yyyy hh:mm:s tt";
            this.dtp_startTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtp_startTime.Location = new System.Drawing.Point(24, 21);
            this.dtp_startTime.Name = "dtp_startTime";
            this.dtp_startTime.Size = new System.Drawing.Size(160, 22);
            this.dtp_startTime.TabIndex = 43;
            //
            // dtp_stopTime
            //
            this.dtp_stopTime.CustomFormat = "MM/dd/yyyy hh:mm:s tt";
            this.dtp_stopTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtp_stopTime.Location = new System.Drawing.Point(191, 21);
            this.dtp_stopTime.Name = "dtp_stopTime";
            this.dtp_stopTime.Size = new System.Drawing.Size(162, 22);
            this.dtp_stopTime.TabIndex = 44;
            //
            // nup_numValues
            //
            this.nup_numValues.Location = new System.Drawing.Point(24, 63);
            this.nup_numValues.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.nup_numValues.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nup_numValues.Name = "nup_numValues";
            this.nup_numValues.Size = new System.Drawing.Size(78, 22);
            this.nup_numValues.TabIndex = 45;
            this.nup_numValues.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            //
            // ckbBrowseAllInitially
            //
            this.ckbBrowseAllInitially.Location = new System.Drawing.Point(151, 16);
            this.ckbBrowseAllInitially.Name = "ckbBrowseAllInitially";
            this.ckbBrowseAllInitially.Size = new System.Drawing.Size(72, 27);
            this.ckbBrowseAllInitially.TabIndex = 20;
            this.ckbBrowseAllInitially.Text = "全初始化";
            //
            // mainSplit
            //
            this.mainSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainSplit.Location = new System.Drawing.Point(0, 0);
            this.mainSplit.Name = "mainSplit";
            //
            // mainSplit.Panel1
            //
            this.mainSplit.Panel1.Controls.Add(this.panel1);
            this.mainSplit.Panel1.Controls.Add(this.groupBox3);
            this.mainSplit.Panel1.Controls.Add(this.flowLayoutPanel1);
            //
            // mainSplit.Panel2
            //
            this.mainSplit.Panel2.Controls.Add(this.secondSplit);
            this.mainSplit.Size = new System.Drawing.Size(744, 370);
            this.mainSplit.SplitterDistance = 226;
            this.mainSplit.SplitterWidth = 6;
            this.mainSplit.TabIndex = 46;
            //
            // panel1
            //
            this.panel1.Controls.Add(this.cbOPCServers);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 49);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(226, 286);
            this.panel1.TabIndex = 2;
            //
            // groupBox3
            //
            this.groupBox3.Controls.Add(this.cbHostName);
            this.groupBox3.Controls.Add(this.ckbBrowseAllInitially);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox3.Location = new System.Drawing.Point(0, 0);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(226, 49);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Host";
            //
            // flowLayoutPanel1
            //
            this.flowLayoutPanel1.Controls.Add(this.btnBrowseOPCServers);
            this.flowLayoutPanel1.Controls.Add(this.btnConnect);
            this.flowLayoutPanel1.Controls.Add(this.btnDisconnect);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 335);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(226, 35);
            this.flowLayoutPanel1.TabIndex = 0;
            //
            // secondSplit
            //
            this.secondSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.secondSplit.Enabled = false;
            this.secondSplit.Location = new System.Drawing.Point(0, 0);
            this.secondSplit.Name = "secondSplit";
            //
            // secondSplit.Panel1
            //
            this.secondSplit.Panel1.Controls.Add(this.panel4);
            this.secondSplit.Panel1.Controls.Add(this.panel3);
            //
            // secondSplit.Panel2
            //
            this.secondSplit.Panel2.Controls.Add(this.tableLayoutPanel1);
            this.secondSplit.Size = new System.Drawing.Size(512, 370);
            this.secondSplit.SplitterDistance = 139;
            this.secondSplit.TabIndex = 0;
            //
            // panel4
            //
            this.panel4.Controls.Add(this.tvServer);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 23);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(139, 347);
            this.panel4.TabIndex = 7;
            //
            // panel3
            //
            this.panel3.Controls.Add(this.btnReset);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(139, 23);
            this.panel3.TabIndex = 6;
            //
            // btnReset
            //
            this.btnReset.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnReset.Location = new System.Drawing.Point(0, 0);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(139, 23);
            this.btnReset.TabIndex = 5;
            this.btnReset.Text = "重新整理";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            //
            // tableLayoutPanel1
            //
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.tbValues, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.groupBox1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.groupBox2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.groupBox4, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 88F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(369, 370);
            this.tableLayoutPanel1.TabIndex = 0;
            //
            // groupBox4
            //
            this.tableLayoutPanel1.SetColumnSpan(this.groupBox4, 2);
            this.groupBox4.Controls.Add(this.dtp_startTime);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.dtp_stopTime);
            this.groupBox4.Controls.Add(this.nup_numValues);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox4.Location = new System.Drawing.Point(3, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(363, 94);
            this.groupBox4.TabIndex = 41;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "時間間隔";
            //
            // panel2
            //
            this.panel2.Controls.Add(this.tbStatus);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 370);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(744, 21);
            this.panel2.TabIndex = 47;
            //
            // mainPanel
            //
            this.mainPanel.Controls.Add(this.mainSplit);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(744, 370);
            this.mainPanel.TabIndex = 48;
            //
            // Form1
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 15);
            this.ClientSize = new System.Drawing.Size(744, 391);
            this.Controls.Add(this.mainPanel);
            this.Controls.Add(this.panel2);
            this.MinimumSize = new System.Drawing.Size(760, 430);
            this.Name = "Form1";
            this.Text = "OPCHDA.NET Synchronous Read Sample";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.timeInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nup_numValues)).EndInit();
            this.mainSplit.Panel1.ResumeLayout(false);
            this.mainSplit.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.mainSplit)).EndInit();
            this.mainSplit.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.secondSplit.Panel1.ResumeLayout(false);
            this.secondSplit.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.secondSplit)).EndInit();
            this.secondSplit.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.mainPanel.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion Windows Form Designer generated code

        private void btn_export_Click(object sender, EventArgs e)
        {
            tbStatus.Text = "開始進行匯出作業....";

            OPCHDAtime startTime = new OPCHDAtime();     // start time
            startTime.Time = Convert.ToDateTime(dtp_startTime.Value).AddHours(-8);
            OPCHDAtime stopTime = new OPCHDAtime();     // stop time
            stopTime.Time = Convert.ToDateTime(dtp_stopTime.Value).AddHours(-8);
            int numValues = Convert.ToInt32(nup_numValues.Value);

            int dtResult = DateTime.Compare(startTime.Time, stopTime.Time);
            if (dtResult == 0)
            {
                tbStatus.Text = "開始時間與結束時間相同";
            }
            else if (dtResult > 0)
            {
                tbStatus.Text = "結束時間不得早於開始時間";
            }
            else if (stopTime.Time > DateTime.Now)
            {
                tbStatus.Text = "結束時間不得晚於當前時間";
            }
            else
            {
                //計算刻度
                //TimeSpan sp = TimeSpan.FromTicks((stopTime.Time.Subtract(startTime.Time).Ticks) / tick);
                //int dateCount = Convert.ToInt32(sp.TotalSeconds);

                //項目清單
                List<TreeNode> nodes = new List<TreeNode>();
                collectChkNodes(ref nodes, tvServer.Nodes);

                //無項目
                if (nodes.Count == 0)
                {
                    tbStatus.Text = "無選取項目";
                    return;
                }
                else
                {
                    tbStatus.Text = "指定檔案路徑";

                    //存檔
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "csv(*.csv) | *.csv | 所有檔案(*.*) | *.* ";
                    saveFileDialog1.Title = "csv匯出";
                    saveFileDialog1.ShowDialog();
                    string strFilePath = saveFileDialog1.FileName.ToString();
                    if (saveFileDialog1.FileName != "")
                    {
                        tbStatus.Text = "csv匯出中，請稍候....";
                        System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                        System.IO.FileInfo FileAttribute = new FileInfo(strFilePath);
                        FileAttribute.Attributes = FileAttributes.Normal;
                        StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);

                        //欄位建立
                        sw.Write("Timestamp,");
                        foreach (TreeNode node in nodes)
                        {
                            sw.Write(node.FullPath.ToString() + ",");
                        }
                        sw.Write("\n");

                        //HDA Read環境建置
                        string[] items = new string[nodes.Count];
                        int[] clientHandles = new int[nodes.Count];
                        int[] serverHandles;
                        int[] err;

                        int idx = 0;
                        foreach (TreeNode tn in nodes)
                        {
                            items[idx] = tn.FullPath;
                            clientHandles[idx] = idx;
                            idx++;
                        }

                        //Handle取得
                        int rtc = OpcSrv.GetItemHandles(items, clientHandles, out serverHandles, out err);
                        if (HRESULTS.Failed(rtc))
                        {
                            Console.WriteLine("Handle取得失敗。錯誤代碼:0x" + rtc.ToString("X"));
                            return;
                        }

                        //開始時間的下一間隔
                        OPCHDAtime tmp_end = new OPCHDAtime();
                        //數值填寫
                        while (startTime.Time <= stopTime.Time)
                        {
                            //讀取Raw Data
                            /*
                             * Tim 修改
                            tmp_end.Time = startTime.Time.AddSeconds(Convert.ToDouble(timeInterval.Value));
                            OPCHDAitem[] values;
                            rtc = OpcSrv.ReadRaw(startTime, tmp_end, numValues, chkBounds.Checked, serverHandles, out values, out err);
                             */

                            OPCHDAitem[] values;
                            tmp_end.Time = startTime.Time.AddSeconds(Convert.ToDouble(timeInterval.Value) * -1);
                            rtc = OpcSrv.ReadRaw(tmp_end, startTime, numValues, chkBounds.Checked, serverHandles, out values, out err);

                            if (HRESULTS.Failed(rtc))
                            {
                                Console.WriteLine("讀取失敗。錯誤代碼:0x" + rtc.ToString("X"));
                            }

                            sw.Write(startTime.Time.AddHours(8).ToString() + ",");
                            idx = 0;
                            foreach (OPCHDAitem opchdaitem in values)
                            {
                                if (opchdaitem.DataValues.Length == 0 || opchdaitem.Client > idx)
                                {
                                    sw.Write("No Value" + ",");
                                }
                                else if (opchdaitem.DataValues[0] != null)
                                {
                                    sw.Write(opchdaitem.DataValues[0].ToString() + ",");
                                }
                                else
                                {
                                    sw.Write("No Value" + ",");
                                }

                                idx++;
                            }

                            sw.Write("\n");
                            startTime.Time = startTime.Time.AddSeconds(Convert.ToDouble(timeInterval.Value));
                        }

                        sw.Close();
                        tbStatus.Text = "匯出完成";
                    }
                    else
                    {
                        tbStatus.Text = "匯出作業取消";
                    }

                    while (startTime.Time <= stopTime.Time)
                    {
                        startTime.Time = startTime.Time.AddSeconds(Convert.ToDouble(timeInterval.Value));
                    }
                }
            }
        }

        //-------------------------------------------------------------------
        // Browse the installed OPC Servers
        private void btnBrowseOPCServers_Click(object sender, System.EventArgs e)
        {
            tbStatus.Text = "正在搜尋所有HDA Server，請稍候...";
            cbOPCServers.Items.Clear();

            OpcHDAServerBrowser SrvList = new OpcHDAServerBrowser(this.cbHostName.Text);
            string[] Servers;
            SrvList.GetServerList(out Servers);

            int hdaSvr_count = 0;
            if (Servers.Length != 0)
            {
                cbOPCServers.Items.AddRange(Servers);
                hdaSvr_count += Servers.Length;
            }
            tbStatus.Text = "HDA Server搜尋結果如下\t個數:" + hdaSvr_count.ToString();
        }

        //--------------------------------------------------------------
        // Connect to the selected OPC Server
        private void btnConnect_Click(object sender, System.EventArgs e)
        {
            tbStatus.Text = "嘗試連線HDA Server";
            tvServer.Nodes.Clear();

            if (cbOPCServers.SelectedIndex == -1)
            {
                tbStatus.Text = "未選擇HDA Server";
            }
            else
            {
                OpcSrv = new OpcHDAServer();
                int rtc = OpcSrv.Connect(cbHostName.Text, cbOPCServers.SelectedItem.ToString());
                if (HRESULTS.Failed(rtc))
                {
                    tbStatus.Text = "連線失敗。請查看結果代碼:0x" + rtc.ToString("X");
                }
                else
                {
                    try
                    {
                        OpcSrv.ShutdownRequested += new ShutdownRequestEventHandler(srv_ShutdownRequested);
                        resetTreeview();
                        //相關控制項保持鎖定
                        ConnectSwitch(false);

                        tbStatus.Text = "已連上" + cbOPCServers.SelectedItem.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        tbStatus.Text = "連線失敗: Kepserver試用過期";
                    }
                }
            }
        }

        //-----------------------------------------------------------------
        // Disconnect from the OPC Server
        private void btnDisconnect_Click(object sender, System.EventArgs e)
        {
            if (OpcSrv.isConnected())
            {
                tvServer.Nodes.Clear();
                tbValues.Text = "";

                //OpcSrv.Disconnect();

                tbStatus.Text = "已中斷與" + cbOPCServers.SelectedItem.ToString() + "連線";

                ConnectSwitch(true);
            }
            else
            {
                tbStatus.Text = "尚未與" + cbOPCServers.SelectedItem.ToString() + "建立連線";
            }
        }

        //--------------------------------------------------------------------
        // Read n Values from the StartTime for the selected item
        private void btnRead_Click(object sender, System.EventArgs e)
        {
            this.tbValues.Text = "";     // clear previous displayd values
            this.tbStatus.Text = "正在讀取項目數值";

            OPCHDAtime startTime = new OPCHDAtime();     // start time
            startTime.Time = Convert.ToDateTime(dtp_startTime.Value).AddHours(-8);
            OPCHDAtime stopTime = new OPCHDAtime();     // stop time
            stopTime.Time = Convert.ToDateTime(dtp_stopTime.Value).AddHours(-8);
            int numValues = Convert.ToInt32(nup_numValues.Value);

            int dtResult = DateTime.Compare(startTime.Time, stopTime.Time);
            if (dtResult == 0)
            {
                tbStatus.Text = "開始時間與結束時間相同";
            }
            else if (dtResult > 0)
            {
                tbStatus.Text = "結束時間不得早於開始時間";
            }
            else if (stopTime.Time > DateTime.Now)
            {
                tbStatus.Text = "結束時間不得晚於當前時間";
            }
            else
            {
                List<TreeNode> nodes = new List<TreeNode>();
                collectChkNodes(ref nodes, tvServer.Nodes);

                if (nodes.Count == 0)
                {
                    tbStatus.Text = "無選取項目";
                    return;
                }
                else
                {
                    string[] items = new string[nodes.Count];
                    int[] clientHandles = new int[nodes.Count];
                    int[] serverHandles;
                    int[] err;

                    int idx = 0;
                    foreach (TreeNode tn in nodes)
                    {
                        items[idx] = tn.FullPath;
                        clientHandles[idx] = idx;
                        idx++;
                    }

                    int rtc = OpcSrv.GetItemHandles(items, clientHandles, out serverHandles, out err);
                    if (HRESULTS.Failed(rtc))
                    {
                        tbStatus.Text = "Handle取得失敗。錯誤代碼:0x" + rtc.ToString("X");
                        return;
                    }

                    OPCHDAitem[] values;
                    rtc = OpcSrv.ReadRaw(startTime, stopTime, numValues, chkBounds.Checked, serverHandles, out values, out err);
                    if (HRESULTS.Failed(rtc))
                    {
                        tbStatus.Text = "讀取失敗。錯誤代碼:0x" + rtc.ToString("X");
                        return;
                    }

                    StringBuilder sb = new StringBuilder();
                    idx = 0;
                    foreach (OPCHDAitem opchdaitem in values)
                    {
                        sb.AppendLine(items[idx] + " :\r\n");

                        for (int v = 0; v < opchdaitem.DataValues.Length; ++v)
                        {
                            if (opchdaitem.DataValues[v] != null)
                            {
                                sb.AppendLine(opchdaitem.TimeStamps[v].ToLocalTime().ToString("MM/dd/yyyy HH:mm:ss.fffffff") + "\t" +
                                opchdaitem.DataValues[v].ToString() + "\t0x" + opchdaitem.Qualities[v].ToString("X"));
                            }
                            else
                            {
                                sb.AppendLine("???\t0x" + opchdaitem.Qualities[v].ToString("X"));
                            }
                        }

                        sb.AppendLine("\n\n");
                        idx++;
                    }
                    tbValues.Text = sb.ToString();
                }
                this.tbStatus.Text = "已完成項目讀取";
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            resetTreeview();
        }

        private void collectChkNodes(ref List<TreeNode> tnc, TreeNodeCollection src)
        {
            foreach (TreeNode tn in src)
            {
                if (tn.Checked && tn.Nodes.Count == 0)
                {
                    tnc.Add(tn);
                }

                if (tn.Nodes.Count > 0)
                {
                    foreach (TreeNode childtn in tn.Nodes)
                    {
                        collectChkNodes(ref tnc, childtn.Nodes);
                    }
                }
            }
        }

        private void ConnectSwitch(bool flag)
        {
            groupBox3.Enabled = flag;
            cbOPCServers.Enabled = flag;
            btnBrowseOPCServers.Enabled = flag;
            btnConnect.Enabled = flag;
            btnDisconnect.Enabled = flag ? false : true;
            secondSplit.Enabled = flag ? false : true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //get now date
            // txt_date_end.Text = thisday.ToLocalTime().ToString("MM/dd/yyyy hh:mm:ss");
            //OpcSrv.ShutdownRequested += new ShutdownRequestEventHandler(srv_ShutdownRequested);
        }

        //------------------------------------------------------------------
        // Browse the OPC Server and display it's items
        private void resetTreeview()
        {
            //tvServer.Nodes.Clear();

            OPCHDABrowser browser;
            int[] err;
            int rtc = OpcSrv.CreateBrowse(null, out browser, out err);

            ShowHDABrowseTree BTree = new ShowHDABrowseTree(browser, tvServer);
            if (this.ckbBrowseAllInitially.Checked)
                BTree.BrowseModeOneLevel = false;

            rtc = BTree.Show();     // display all branches and items
        }

        private void srv_ShutdownRequested(object sender, ShutdownRequestEventArgs e)
        {
            tbStatus.Text = "(Server Shutdown Callback)" + "已中斷與HDA Server連線，原因: " + e.shutdownReason;
            ConnectSwitch(true);
        }
    }
}