using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

using SerialComm;
using mshtml;

using SharpTerminal;


// TODO: Allow color customization
namespace SharpTerminal
{
	/// <summary>
	/// Summary description for MainForm.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.ClearButton = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.HexMode = new System.Windows.Forms.RadioButton();
			this.TextMode = new System.Windows.Forms.RadioButton();
			this.SendButton = new System.Windows.Forms.Button();
			this.SendText = new System.Windows.Forms.TextBox();
			this.ComPort = new System.Windows.Forms.TextBox();
			this.DisconnectButton = new System.Windows.Forms.Button();
			this.CtsLabel = new System.Windows.Forms.Label();
			this.RtsButton = new System.Windows.Forms.Button();
			this.ConnectButton = new System.Windows.Forms.Button();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.label6 = new System.Windows.Forms.Label();
			this.SaveDefaultSettingsButton = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.LineEndChars = new System.Windows.Forms.ComboBox();
			this.AutoAddLineEnd = new System.Windows.Forms.CheckBox();
			this.ShowCRLF = new System.Windows.Forms.CheckBox();
			this.ShowLF = new System.Windows.Forms.CheckBox();
			this.ShowCR = new System.Windows.Forms.CheckBox();
			this.FlowControlList = new System.Windows.Forms.ComboBox();
			this.label5 = new System.Windows.Forms.Label();
			this.StopBitsList = new System.Windows.Forms.ComboBox();
			this.label4 = new System.Windows.Forms.Label();
			this.ParityList = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.DataSizeList = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.BaudRateList = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.EchoLocal = new System.Windows.Forms.CheckBox();
			this.ShowConnections = new System.Windows.Forms.CheckBox();
			this.ShowCtsTransitions = new System.Windows.Forms.CheckBox();
			this.ShowRtsTransitions = new System.Windows.Forms.CheckBox();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.FileMenu = new System.Windows.Forms.MenuItem();
			this.FileMenu_OpenTranscript = new System.Windows.Forms.MenuItem();
			this.FileMenu_SaveTranscript = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.FileMenu_Exit = new System.Windows.Forms.MenuItem();
			this.HelpMenu = new System.Windows.Forms.MenuItem();
			this.HelpMenu_About = new System.Windows.Forms.MenuItem();
			this.webMain = new AxSHDocVw.AxWebBrowser();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.webMain)).BeginInit();
			this.SuspendLayout();
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControl1.Location = new System.Drawing.Point(0, 0);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(664, 465);
			this.tabControl1.TabIndex = 0;
			this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.webMain);
			this.tabPage1.Controls.Add(this.ClearButton);
			this.tabPage1.Controls.Add(this.groupBox1);
			this.tabPage1.Controls.Add(this.SendButton);
			this.tabPage1.Controls.Add(this.SendText);
			this.tabPage1.Controls.Add(this.ComPort);
			this.tabPage1.Controls.Add(this.DisconnectButton);
			this.tabPage1.Controls.Add(this.CtsLabel);
			this.tabPage1.Controls.Add(this.RtsButton);
			this.tabPage1.Controls.Add(this.ConnectButton);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(616, 439);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Communication";
			// 
			// ClearButton
			// 
			this.ClearButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ClearButton.Location = new System.Drawing.Point(433, 16);
			this.ClearButton.Name = "ClearButton";
			this.ClearButton.TabIndex = 23;
			this.ClearButton.Text = "Clear";
			this.ClearButton.Click += new System.EventHandler(this.ClearButton_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.HexMode);
			this.groupBox1.Controls.Add(this.TextMode);
			this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBox1.Location = new System.Drawing.Point(293, 7);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(129, 35);
			this.groupBox1.TabIndex = 22;
			this.groupBox1.TabStop = false;
			// 
			// HexMode
			// 
			this.HexMode.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.HexMode.Location = new System.Drawing.Point(72, 14);
			this.HexMode.Name = "HexMode";
			this.HexMode.Size = new System.Drawing.Size(48, 16);
			this.HexMode.TabIndex = 1;
			this.HexMode.Text = "Hex";
			this.HexMode.CheckedChanged += new System.EventHandler(this.HexMode_CheckedChanged);
			// 
			// TextMode
			// 
			this.TextMode.Checked = true;
			this.TextMode.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.TextMode.Location = new System.Drawing.Point(7, 14);
			this.TextMode.Name = "TextMode";
			this.TextMode.Size = new System.Drawing.Size(55, 16);
			this.TextMode.TabIndex = 0;
			this.TextMode.TabStop = true;
			this.TextMode.Text = "Text";
			this.TextMode.CheckedChanged += new System.EventHandler(this.TextMode_CheckedChanged);
			// 
			// SendButton
			// 
			this.SendButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.SendButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.SendButton.Location = new System.Drawing.Point(538, 411);
			this.SendButton.Name = "SendButton";
			this.SendButton.Size = new System.Drawing.Size(75, 22);
			this.SendButton.TabIndex = 20;
			this.SendButton.Text = "Send";
			this.SendButton.Click += new System.EventHandler(this.SendButton_Click);
			// 
			// SendText
			// 
			this.SendText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.SendText.Location = new System.Drawing.Point(7, 412);
			this.SendText.Name = "SendText";
			this.SendText.Size = new System.Drawing.Size(524, 20);
			this.SendText.TabIndex = 19;
			this.SendText.Text = "";
			this.SendText.KeyUp += new System.Windows.Forms.KeyEventHandler(this.SendText_KeyUp);
			// 
			// ComPort
			// 
			this.ComPort.Location = new System.Drawing.Point(8, 16);
			this.ComPort.Name = "ComPort";
			this.ComPort.TabIndex = 13;
			this.ComPort.Text = "COM1";
			// 
			// DisconnectButton
			// 
			this.DisconnectButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.DisconnectButton.Location = new System.Drawing.Point(208, 16);
			this.DisconnectButton.Name = "DisconnectButton";
			this.DisconnectButton.Size = new System.Drawing.Size(80, 23);
			this.DisconnectButton.TabIndex = 18;
			this.DisconnectButton.Text = "Disconnect";
			this.DisconnectButton.Click += new System.EventHandler(this.DisconnectButton_Click);
			// 
			// CtsLabel
			// 
			this.CtsLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.CtsLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.CtsLabel.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.CtsLabel.Location = new System.Drawing.Point(525, 18);
			this.CtsLabel.Name = "CtsLabel";
			this.CtsLabel.Size = new System.Drawing.Size(40, 21);
			this.CtsLabel.TabIndex = 16;
			this.CtsLabel.Text = "CTS";
			this.CtsLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// RtsButton
			// 
			this.RtsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.RtsButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.RtsButton.Location = new System.Drawing.Point(571, 18);
			this.RtsButton.Name = "RtsButton";
			this.RtsButton.Size = new System.Drawing.Size(40, 21);
			this.RtsButton.TabIndex = 15;
			this.RtsButton.Text = "RTS";
			this.RtsButton.Click += new System.EventHandler(this.RtsButton_Click);
			// 
			// ConnectButton
			// 
			this.ConnectButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ConnectButton.Location = new System.Drawing.Point(120, 16);
			this.ConnectButton.Name = "ConnectButton";
			this.ConnectButton.TabIndex = 14;
			this.ConnectButton.Text = "Connect";
			this.ConnectButton.Click += new System.EventHandler(this.ConnectButton_Click);
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.label6);
			this.tabPage2.Controls.Add(this.SaveDefaultSettingsButton);
			this.tabPage2.Controls.Add(this.groupBox2);
			this.tabPage2.Controls.Add(this.FlowControlList);
			this.tabPage2.Controls.Add(this.label5);
			this.tabPage2.Controls.Add(this.StopBitsList);
			this.tabPage2.Controls.Add(this.label4);
			this.tabPage2.Controls.Add(this.ParityList);
			this.tabPage2.Controls.Add(this.label3);
			this.tabPage2.Controls.Add(this.DataSizeList);
			this.tabPage2.Controls.Add(this.label2);
			this.tabPage2.Controls.Add(this.BaudRateList);
			this.tabPage2.Controls.Add(this.label1);
			this.tabPage2.Controls.Add(this.groupBox3);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(656, 439);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Configuration";
			// 
			// label6
			// 
			this.label6.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label6.Location = new System.Drawing.Point(448, 512);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(200, 23);
			this.label6.TabIndex = 13;
			this.label6.Text = "Make these settings my default:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// SaveDefaultSettingsButton
			// 
			this.SaveDefaultSettingsButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.SaveDefaultSettingsButton.Location = new System.Drawing.Point(648, 512);
			this.SaveDefaultSettingsButton.Name = "SaveDefaultSettingsButton";
			this.SaveDefaultSettingsButton.TabIndex = 12;
			this.SaveDefaultSettingsButton.Text = "Save";
			this.SaveDefaultSettingsButton.Click += new System.EventHandler(this.SaveDefaultSettingsButton_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.LineEndChars);
			this.groupBox2.Controls.Add(this.AutoAddLineEnd);
			this.groupBox2.Controls.Add(this.ShowCRLF);
			this.groupBox2.Controls.Add(this.ShowLF);
			this.groupBox2.Controls.Add(this.ShowCR);
			this.groupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBox2.Location = new System.Drawing.Point(232, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(184, 168);
			this.groupBox2.TabIndex = 10;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Special Characters";
			// 
			// LineEndChars
			// 
			this.LineEndChars.Items.AddRange(new object[] {
																											"0x0D",
																											"0x0A",
																											"0x0D0A"});
			this.LineEndChars.Location = new System.Drawing.Point(32, 136);
			this.LineEndChars.Name = "LineEndChars";
			this.LineEndChars.Size = new System.Drawing.Size(144, 21);
			this.LineEndChars.TabIndex = 4;
			// 
			// AutoAddLineEnd
			// 
			this.AutoAddLineEnd.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
			this.AutoAddLineEnd.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.AutoAddLineEnd.Location = new System.Drawing.Point(8, 96);
			this.AutoAddLineEnd.Name = "AutoAddLineEnd";
			this.AutoAddLineEnd.Size = new System.Drawing.Size(168, 40);
			this.AutoAddLineEnd.TabIndex = 3;
			this.AutoAddLineEnd.Text = "Automatically append line end to messages:";
			this.AutoAddLineEnd.TextAlign = System.Drawing.ContentAlignment.TopLeft;
			// 
			// ShowCRLF
			// 
			this.ShowCRLF.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowCRLF.Location = new System.Drawing.Point(8, 72);
			this.ShowCRLF.Name = "ShowCRLF";
			this.ShowCRLF.Size = new System.Drawing.Size(168, 24);
			this.ShowCRLF.TabIndex = 2;
			this.ShowCRLF.Text = "Display incoming CRLF";
			// 
			// ShowLF
			// 
			this.ShowLF.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowLF.Location = new System.Drawing.Point(8, 48);
			this.ShowLF.Name = "ShowLF";
			this.ShowLF.Size = new System.Drawing.Size(168, 24);
			this.ShowLF.TabIndex = 1;
			this.ShowLF.Text = "Display incoming LF";
			// 
			// ShowCR
			// 
			this.ShowCR.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowCR.Location = new System.Drawing.Point(8, 24);
			this.ShowCR.Name = "ShowCR";
			this.ShowCR.Size = new System.Drawing.Size(168, 24);
			this.ShowCR.TabIndex = 0;
			this.ShowCR.Text = "Display incoming CR";
			// 
			// FlowControlList
			// 
			this.FlowControlList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.FlowControlList.Items.AddRange(new object[] {
																												 "None",
																												 "Hardware",
																												 "XOn/XOff"});
			this.FlowControlList.Location = new System.Drawing.Point(104, 136);
			this.FlowControlList.Name = "FlowControlList";
			this.FlowControlList.Size = new System.Drawing.Size(121, 21);
			this.FlowControlList.TabIndex = 9;
			this.FlowControlList.SelectedIndexChanged += new System.EventHandler(this.FlowControlList_SelectedIndexChanged);
			// 
			// label5
			// 
			this.label5.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label5.Location = new System.Drawing.Point(8, 136);
			this.label5.Name = "label5";
			this.label5.TabIndex = 8;
			this.label5.Text = "Flow Control";
			// 
			// StopBitsList
			// 
			this.StopBitsList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.StopBitsList.Items.AddRange(new object[] {
																											"1",
																											"1.5",
																											"2"});
			this.StopBitsList.Location = new System.Drawing.Point(104, 104);
			this.StopBitsList.Name = "StopBitsList";
			this.StopBitsList.Size = new System.Drawing.Size(121, 21);
			this.StopBitsList.TabIndex = 7;
			this.StopBitsList.SelectedIndexChanged += new System.EventHandler(this.StopBitsList_SelectedIndexChanged);
			// 
			// label4
			// 
			this.label4.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label4.Location = new System.Drawing.Point(8, 104);
			this.label4.Name = "label4";
			this.label4.TabIndex = 6;
			this.label4.Text = "Stop Bits";
			// 
			// ParityList
			// 
			this.ParityList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ParityList.Items.AddRange(new object[] {
																										"None",
																										"Even",
																										"Odd"});
			this.ParityList.Location = new System.Drawing.Point(104, 72);
			this.ParityList.Name = "ParityList";
			this.ParityList.Size = new System.Drawing.Size(121, 21);
			this.ParityList.TabIndex = 5;
			this.ParityList.SelectedIndexChanged += new System.EventHandler(this.ParityList_SelectedIndexChanged);
			// 
			// label3
			// 
			this.label3.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label3.Location = new System.Drawing.Point(8, 72);
			this.label3.Name = "label3";
			this.label3.TabIndex = 4;
			this.label3.Text = "Parity";
			// 
			// DataSizeList
			// 
			this.DataSizeList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.DataSizeList.Items.AddRange(new object[] {
																											"4",
																											"5",
																											"6",
																											"7",
																											"8"});
			this.DataSizeList.Location = new System.Drawing.Point(104, 40);
			this.DataSizeList.Name = "DataSizeList";
			this.DataSizeList.Size = new System.Drawing.Size(121, 21);
			this.DataSizeList.TabIndex = 3;
			this.DataSizeList.SelectedIndexChanged += new System.EventHandler(this.DataSizeList_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.TabIndex = 2;
			this.label2.Text = "Data Bits";
			// 
			// BaudRateList
			// 
			this.BaudRateList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.BaudRateList.Items.AddRange(new object[] {
																											"1200",
																											"2400",
																											"4800",
																											"9600",
																											"19200",
																											"38400",
																											"57600",
																											"115200"});
			this.BaudRateList.Location = new System.Drawing.Point(104, 8);
			this.BaudRateList.Name = "BaudRateList";
			this.BaudRateList.Size = new System.Drawing.Size(121, 21);
			this.BaudRateList.TabIndex = 1;
			this.BaudRateList.SelectedIndexChanged += new System.EventHandler(this.BaudRateList_SelectedIndexChanged);
			// 
			// label1
			// 
			this.label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.TabIndex = 0;
			this.label1.Text = "Baud Rate";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.EchoLocal);
			this.groupBox3.Controls.Add(this.ShowConnections);
			this.groupBox3.Controls.Add(this.ShowCtsTransitions);
			this.groupBox3.Controls.Add(this.ShowRtsTransitions);
			this.groupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBox3.Location = new System.Drawing.Point(424, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(232, 168);
			this.groupBox3.TabIndex = 11;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Transcript Options";
			// 
			// EchoLocal
			// 
			this.EchoLocal.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.EchoLocal.Location = new System.Drawing.Point(8, 96);
			this.EchoLocal.Name = "EchoLocal";
			this.EchoLocal.Size = new System.Drawing.Size(216, 24);
			this.EchoLocal.TabIndex = 3;
			this.EchoLocal.Text = "Echo Sent Messages";
			// 
			// ShowConnections
			// 
			this.ShowConnections.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowConnections.Location = new System.Drawing.Point(8, 72);
			this.ShowConnections.Name = "ShowConnections";
			this.ShowConnections.Size = new System.Drawing.Size(216, 24);
			this.ShowConnections.TabIndex = 2;
			this.ShowConnections.Text = "Include Connections and Disconnections";
			// 
			// ShowCtsTransitions
			// 
			this.ShowCtsTransitions.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowCtsTransitions.Location = new System.Drawing.Point(8, 48);
			this.ShowCtsTransitions.Name = "ShowCtsTransitions";
			this.ShowCtsTransitions.Size = new System.Drawing.Size(216, 24);
			this.ShowCtsTransitions.TabIndex = 1;
			this.ShowCtsTransitions.Text = "Include CTS line transitions";
			// 
			// ShowRtsTransitions
			// 
			this.ShowRtsTransitions.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.ShowRtsTransitions.Location = new System.Drawing.Point(8, 24);
			this.ShowRtsTransitions.Name = "ShowRtsTransitions";
			this.ShowRtsTransitions.Size = new System.Drawing.Size(216, 24);
			this.ShowRtsTransitions.TabIndex = 0;
			this.ShowRtsTransitions.Text = "Include RTS line transitions";
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.FileMenu,
																																							this.HelpMenu});
			// 
			// FileMenu
			// 
			this.FileMenu.Index = 0;
			this.FileMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																						 this.FileMenu_OpenTranscript,
																																						 this.FileMenu_SaveTranscript,
																																						 this.menuItem3,
																																						 this.FileMenu_Exit});
			this.FileMenu.Text = "&File";
			// 
			// FileMenu_OpenTranscript
			// 
			this.FileMenu_OpenTranscript.Index = 0;
			this.FileMenu_OpenTranscript.Shortcut = System.Windows.Forms.Shortcut.CtrlO;
			this.FileMenu_OpenTranscript.Text = "&Open Transcript...";
			this.FileMenu_OpenTranscript.Click += new System.EventHandler(this.FileMenu_OpenTranscript_Click);
			// 
			// FileMenu_SaveTranscript
			// 
			this.FileMenu_SaveTranscript.Index = 1;
			this.FileMenu_SaveTranscript.Shortcut = System.Windows.Forms.Shortcut.CtrlS;
			this.FileMenu_SaveTranscript.Text = "&Save Transcript...";
			this.FileMenu_SaveTranscript.Click += new System.EventHandler(this.FileMenu_SaveTranscript_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 2;
			this.menuItem3.Text = "-";
			// 
			// FileMenu_Exit
			// 
			this.FileMenu_Exit.Index = 3;
			this.FileMenu_Exit.Shortcut = System.Windows.Forms.Shortcut.AltF4;
			this.FileMenu_Exit.Text = "E&xit";
			this.FileMenu_Exit.Click += new System.EventHandler(this.FileMenu_Exit_Click);
			// 
			// HelpMenu
			// 
			this.HelpMenu.Index = 1;
			this.HelpMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																						 this.HelpMenu_About});
			this.HelpMenu.Text = "&Help";
			// 
			// HelpMenu_About
			// 
			this.HelpMenu_About.Index = 0;
			this.HelpMenu_About.Text = "&About";
			this.HelpMenu_About.Click += new System.EventHandler(this.HelpMenu_About_Click);
			// 
			// webMain
			// 
			this.webMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.webMain.ContainingControl = this;
			this.webMain.Enabled = true;
			this.webMain.Location = new System.Drawing.Point(8, 48);
			this.webMain.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("webMain.OcxState")));
			this.webMain.Size = new System.Drawing.Size(600, 360);
			this.webMain.TabIndex = 24;
			// 
			// MainForm
			// 
			this.AcceptButton = this.SendButton;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(664, 465);
			this.Controls.Add(this.tabControl1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu1;
			this.MinimumSize = new System.Drawing.Size(672, 520);
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "SharpTerminal";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MainForm_Closing);
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.webMain)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.EnableVisualStyles();
			Application.DoEvents();
			Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
			Application.Run(new MainForm());
		}

		#region Window Controls
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton HexMode;
		private System.Windows.Forms.RadioButton TextMode;
		private System.Windows.Forms.Button SendButton;
		private System.Windows.Forms.TextBox SendText;
		private System.Windows.Forms.TextBox ComPort;
		private System.Windows.Forms.Button DisconnectButton;
		private System.Windows.Forms.Label CtsLabel;
		private System.Windows.Forms.Button RtsButton;
		private System.Windows.Forms.Button ConnectButton;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.ComboBox BaudRateList;
		private System.Windows.Forms.ComboBox DataSizeList;
		private System.Windows.Forms.ComboBox ParityList;
		private System.Windows.Forms.ComboBox StopBitsList;
		private System.Windows.Forms.ComboBox FlowControlList;
		private System.Windows.Forms.CheckBox ShowCR;
		private System.Windows.Forms.CheckBox ShowLF;
		private System.Windows.Forms.CheckBox ShowCRLF;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem FileMenu;
		private System.Windows.Forms.MenuItem FileMenu_SaveTranscript;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem FileMenu_Exit;
		private System.Windows.Forms.MenuItem HelpMenu;
		private System.Windows.Forms.MenuItem HelpMenu_About;
		private System.Windows.Forms.MenuItem FileMenu_OpenTranscript;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.CheckBox ShowCtsTransitions;
		private System.Windows.Forms.CheckBox ShowRtsTransitions;
		private System.Windows.Forms.CheckBox ShowConnections;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button SaveDefaultSettingsButton;
		private System.Windows.Forms.Button ClearButton;
		#endregion

		private SerialPort _SerialPort;
		private HTMLDocument _Document;
		private IHTMLElement _Body;
		private enum Modes
		{
			Hexadecimal,
			Text
		}
		private Modes _DisplayMode = Modes.Text;
		private bool _Starting = false;
		private bool _CommSettingsChanged = false;
		private string _LastFolder = "";
		private ArrayList _CommandBuffer = new ArrayList();
		private bool _ScanningBuffer = false;
		private int _CurrentBufferIndex = 0;
		private System.Windows.Forms.CheckBox AutoAddLineEnd;
		private System.Windows.Forms.ComboBox LineEndChars;
		private System.Windows.Forms.CheckBox EchoLocal;
		private ArrayList _Transcript = new ArrayList();
		private AxSHDocVw.AxWebBrowser webMain;
		private Encoding _Encoding = Encoding.Default;

		private void MainForm_Load(object sender, System.EventArgs e)
		{
			_Starting = true;
			_SerialPort = new SerialPort();
			LoadDefaultSettings();

			ClearDocument();
			_Document = webMain.Document as HTMLDocument;
			_Body = _Document.body;

			SendButton.Enabled = false;
			DisconnectButton.Enabled = false;
			RtsButton.Enabled = false;
			_Starting = false;
		}

		private void ClearDocument()
		{
			// Get a new blank document
			object o = System.Reflection.Missing.Value;
			webMain.Navigate("about:blank", ref o, ref o, ref o, ref o);
			// Waiting for it to finish loading...
			while (webMain.Document == null || (webMain.Document as HTMLDocument).body == null) Application.DoEvents();
		}

		private void ConnectButton_Click(object sender, System.EventArgs e)
		{
			if (ComPort.Text.Length < 1)
			{
				MessageBox.Show(this, "You must enter a COM port name.");
				return;
			}

			try
			{
				_SerialPort.ComPort = ComPort.Text;
				_SerialPort.BaudRate = (SerialPort.BaudRates)uint.Parse(BaudRateList.Text);
				_SerialPort.DataBits = (SerialPort.DataSizes)uint.Parse(DataSizeList.Text);
				_SerialPort.Parity = (SerialPort.Parities)(uint)ParityList.SelectedIndex;
				_SerialPort.StopBits = (SerialPort.StopSizes)(uint)StopBitsList.SelectedIndex;
				_SerialPort.FlowControl = (SerialPort.FlowControlTypes)(uint)FlowControlList.SelectedIndex;
				_SerialPort.Connect();
				if (ShowConnections.Checked)
				{
					WriteControlMessage(string.Format("C{0}", _SerialPort.ComPort));
				}
				_SerialPort.StatusUpdate += new StatusUpdateEventHandler(_SerialPort_StatusUpdate);
				_SerialPort.Receive += new ReceiveEventHandler(_SerialPort_Receive);
				_SerialPort.WriteComplete += new WriteEventHandler(_SerialPort_WriteComplete);
				_SerialPort.Error += new SerialComm.ErrorEventHandler(_SerialPort_Error);
				SendButton.Enabled = true;
				DisconnectButton.Enabled = true;
				RtsButton.Enabled = true;
				RtsButton.BackColor = (_SerialPort.RtsState ? Color.Lime : Color.LightGray);
			}
			catch (Exception ex)
			{
				MessageBox.Show(this, string.Format("Failed to open {0}: {1}", ComPort.Text, ex.Message));
			}
		}

		private void DisconnectButton_Click(object sender, System.EventArgs e)
		{
			_SerialPort.Disconnect();
			_SerialPort.StatusUpdate -= new StatusUpdateEventHandler(_SerialPort_StatusUpdate);
			_SerialPort.Receive -= new ReceiveEventHandler(_SerialPort_Receive);
			_SerialPort.WriteComplete -= new WriteEventHandler(_SerialPort_WriteComplete);
			_SerialPort.Error -= new SerialComm.ErrorEventHandler(_SerialPort_Error);
			if (ShowConnections.Checked)
			{
				WriteControlMessage(string.Format("D{0}", _SerialPort.ComPort));
			}
			SendButton.Enabled = false;
			DisconnectButton.Enabled = false;
			RtsButton.Enabled = false;
		}

		private void RtsButton_Click(object sender, System.EventArgs e)
		{
			if (ShowRtsTransitions.Checked)
			{
				_SerialPort.RtsState = !_SerialPort.RtsState;
				RtsButton.BackColor = (_SerialPort.RtsState ? Color.Lime : Color.LightGray);
				WriteControlMessage(string.Format("RTS{0}", (_SerialPort.RtsState ? "H" : "L")));
			}
		}

		private void _SerialPort_StatusUpdate(object sender, StatusUpdateEventArgs e)
		{
			if (ShowCtsTransitions.Checked)
			{
				CtsLabel.BackColor = (e.CTS ? Color.Lime : Color.Gray);
				WriteControlMessage(string.Format("CTS{0}", (e.CTS ? "H" : "L")));
			}
		}

		private void _SerialPort_Receive(object sender, ReceiveEventArgs e)
		{
			WriteBytes(e.Data, Directions.Read);
			_Transcript.Add(new Message(DateTime.Now, Directions.Read, e.Data));
		}

		private void _SerialPort_WriteComplete(object sender, WriteEventArgs e)
		{
			if (EchoLocal.Checked)
			{
				WriteBytes(e.Data, Directions.Write);
				_Transcript.Add(new Message(DateTime.Now, Directions.Write, e.Data));
			}
		}

		private void _SerialPort_Error(object sender, SerialComm.ErrorEventArgs e)
		{
			MessageBox.Show("The serial port encountered an error: " + e.Message);
		}
		
		private void WriteControlMessage(string message)
		{
			byte[] data = Encoding.UTF8.GetBytes(message);
			WriteBytes(data, Directions.Control);
			_Transcript.Add(new Message(DateTime.Now, Directions.Control, data));
		}
	
		private void WriteBytes(byte[] data, Directions d)
		{
			if (d == Directions.Control)
			{
				_Body.innerHTML += string.Format("<span style=\"color: blue; word-wrap: break-word;\">{0}</span>", Encoding.UTF8.GetString(data));
				return;
			}

			string color = (d == Directions.Read ? "green" : "black");

			StringBuilder sb = new StringBuilder(string.Format("<span style=\"color: {0}; word-wrap: break-word;\">", color));
			if (_DisplayMode == Modes.Text)
			{
				foreach (byte b in data)
				{
					if (!ShowCR.Checked && b == (byte)'\r')
					{
						sb.Append("<br>");
					}
					else if (!ShowLF.Checked && b == (byte)'\n')
					{
						sb.Append("<br>");
					}
					else if (b < 0x20 || b > 0x7e)
					{
						if (b < 16)
						{
							sb.Append("<sup>");
							sb.Append("0");
							sb.Append("</sup>");
							sb.Append("<sub>");
							sb.Append(b.ToString("x").Substring(0, 1));
							sb.Append("</sub>");
						}
						else
						{
							sb.Append("<sup>");
							sb.Append(b.ToString("x").Substring(0, 1));
							sb.Append("</sup>");
							sb.Append("<sub>");
							sb.Append(b.ToString("x").Substring(1, 1));
							sb.Append("</sub>");
						}
					}
					else
					{
						// Properly handle < in input
						if (b == 0x3C)
						{
							sb.Append("&lt;");
						}
						else
						{
							sb.Append(_Encoding.GetString(new byte[] {b}));
						}
					}
				}
			}
			else
			{
				foreach (byte b in data)
				{
					if (!ShowCR.Checked && b == (byte)'\r')
					{
						sb.Append("<br>");
					}
					else if (!ShowLF.Checked && b == (byte)'\n')
					{
						sb.Append("<br>");
					}
					else if (b < 16)
					{
						sb.Append("<sup>");
						sb.Append("0");
						sb.Append("</sup>");
						sb.Append("<sub>");
						sb.Append(b.ToString("x").Substring(0, 1));
						sb.Append("</sub>");
					}
					else
					{
						sb.Append("<sup>");
						sb.Append(b.ToString("x").Substring(0, 1));
						sb.Append("</sup>");
						sb.Append("<sub>");
						sb.Append(b.ToString("x").Substring(1, 1));
						sb.Append("</sub>");
					}
				}
			}
			sb.Append("</span>");
			_Body.innerHTML += sb.ToString();
		}

		private void SendButton_Click(object sender, System.EventArgs e)
		{
			_ScanningBuffer = false;
			_CommandBuffer.Add(SendText.Text);
			_SerialPort.Write(ConvertIfHex(SendText.Text));
			if (AutoAddLineEnd.Checked && LineEndChars.Text.Length > 0)
			{
				_SerialPort.Write(ConvertIfHex(LineEndChars.Text));
			}
			SendText.Text = "";
		}

		private byte[] ConvertIfHex(string message)
		{

			if (message.StartsWith("0x"))
			{
				// Interpret as hex
				string hex = message.Substring(2);
				if (hex.Length % 2 != 0)
				{
					throw new ArgumentException("You must enter an even number of characters (2 per byte) to send hex data.");
				}
				if (hex.IndexOf(" ") >= 0)
				{
					hex = hex.Replace(" ", "");
				}
				byte[] output = new byte[hex.Length / 2];
				for (int i = 0; i < hex.Length; i += 2)
				{
					output[i / 2] = byte.Parse(hex.Substring(i, 2), System.Globalization.NumberStyles.HexNumber);
				}
				return output;
			}
			else
			{
				// Straight up text
				return _Encoding.GetBytes(message);
			}
		}

		private void SendText_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (_CommandBuffer.Count > 0)
			{
				if (e.KeyCode == Keys.Up)
				{
					if (!_ScanningBuffer)
					{
						_ScanningBuffer = true;
						_CurrentBufferIndex = _CommandBuffer.Count - 1;
					}
					else
					{
						if (_CurrentBufferIndex > 0)
						{
							--_CurrentBufferIndex;
						}
						else
						{
							MessageBeep(0);
						}
					}
					SendText.Text = (string)_CommandBuffer[_CurrentBufferIndex];
				}
				else if (e.KeyCode == Keys.Down && _ScanningBuffer)
				{
					if (_CurrentBufferIndex < _CommandBuffer.Count - 1)
					{
						--_CurrentBufferIndex;
					}
					else
					{
						MessageBeep(0);
					}
					SendText.Text = (string)_CommandBuffer[_CurrentBufferIndex];
				}
			}
		}

		private void ReDisplay()
		{
			// Redisplay the Transcript in the new mode
			_Body.innerHTML = "";
			foreach (Message m in _Transcript)
			{
				WriteBytes(m.Data, m.Direction);
			}
		}

		private void MainForm_Closing(object sender, CancelEventArgs e)
		{
			if (_SerialPort.Connected)
			{
				DialogResult result = MessageBox.Show(this, "You are currently connected.  Are you sure you want to disconnect and close this application?", "Disconnect?", MessageBoxButtons.YesNo);
				if (result == DialogResult.No)
				{
					e.Cancel = true;
					return;
				}
				_SerialPort.Disconnect();
			}
		}

		private void TextMode_CheckedChanged(object sender, System.EventArgs e)
		{
			if (TextMode.Checked)
			{
				_DisplayMode = Modes.Text;
				if (!_Starting)
				{
					ReDisplay();
				}
			}
		}

		private void HexMode_CheckedChanged(object sender, System.EventArgs e)
		{
			if (HexMode.Checked)
			{
				_DisplayMode = Modes.Hexadecimal;
				if (!_Starting)
				{
					ReDisplay();
				}
			}
		}

		private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
		{
			// Package up the error nice and neat
			StringBuilder error = new StringBuilder();
			error.AppendFormat("Date/Time: {0}\n", DateTime.Now.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"));
			error.AppendFormat("Exception Type: {0}\n", e.Exception.GetType().FullName);
			error.AppendFormat("Exception Message: {0}\n", e.Exception.Message);
			error.AppendFormat("Exception Source: {0}\n", e.Exception.Source);
			error.AppendFormat("Exception Target Site: {0}\n", e.Exception.TargetSite);
			error.AppendFormat("Exception Stack Trace: {0}\n", e.Exception.StackTrace);
			Exception i = e.Exception.InnerException;
			string indent = "";
			while (i != null)
			{
				indent += " -> ";
				error.AppendFormat("{0}Inner Exception Type: {1}\n", indent, e.Exception.GetType().FullName);
				error.AppendFormat("{0}Inner Exception Message: {1}\n", indent, e.Exception.Message);
				error.AppendFormat("{0}Inner Exception Source: {1}\n", indent, e.Exception.Source);
				error.AppendFormat("{0}Inner Exception Target Site: {1}\n", indent, e.Exception.TargetSite);
				i = i.InnerException;
			}

			MessageBox.Show("Unhandled exception: " + e.Exception.ToString());
		}

		private void tabControl1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (tabControl1.SelectedIndex == 1)
			{
				// Refresh config info
				BaudRateList.SelectedIndex = BaudRateList.FindStringExact(((uint)(_SerialPort.BaudRate)).ToString());
				DataSizeList.SelectedIndex = DataSizeList.FindStringExact(((uint)(_SerialPort.DataBits)).ToString());
				ParityList.SelectedIndex = ParityList.FindStringExact(_SerialPort.Parity.ToString());
				StopBitsList.SelectedIndex = Convert.ToInt32((uint)_SerialPort.StopBits);
				FlowControlList.SelectedIndex = Convert.ToInt32((uint)_SerialPort.FlowControl);
				_CommSettingsChanged = false;
			}
			else
			{
				DialogResult result = DialogResult.No;
				if (_SerialPort.Connected && _CommSettingsChanged)
				{
					result = MessageBox.Show(this, "You have changed the communication settings while a port is open.  " +
																				 "Do you wish to close the port and reopen it with the new settings?  " +
																				 "If you pick No, the new settings will not be applied until you close " +
																				 "and reopen the port manually.", "Reset COM Port?", MessageBoxButtons.YesNo);
					if (result == DialogResult.Yes)
					{
						DisconnectButton_Click(null, null);
					}
				}
				_SerialPort.BaudRate = (SerialPort.BaudRates)uint.Parse(BaudRateList.Text);
				_SerialPort.DataBits = (SerialPort.DataSizes)uint.Parse(BaudRateList.Text);
				_SerialPort.Parity = (SerialPort.Parities)ParityList.SelectedIndex;
				_SerialPort.StopBits = (SerialPort.StopSizes)StopBitsList.SelectedIndex;
				_SerialPort.FlowControl = (SerialPort.FlowControlTypes)FlowControlList.SelectedIndex;
				if (result == DialogResult.Yes)
				{
					ConnectButton_Click(null, null);
				}
			}
		}

		private void FileMenu_Exit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void FileMenu_SaveTranscript_Click(object sender, System.EventArgs e)
		{
			SaveFileDialog dialog = new SaveFileDialog();
			dialog.AddExtension = true;
			dialog.CheckFileExists = false;
			dialog.CheckPathExists = true;
			dialog.DefaultExt = "stt";
			dialog.FileName = "SessionTranscript";
			dialog.Filter = "SharpTerminal Transcripts (*.stt)|*.stt|Comma-Separated Value Files (*.csv)|*.csv|" + 
											"HTML Files (*.htm;*.html)|*.htm;*.html|All Files (*.*)|*.*";
			dialog.FilterIndex = 1;
			if (_LastFolder.Length == 0)
			{
				dialog.InitialDirectory = _LastFolder;
			}
			else
			{
				dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			}
			dialog.OverwritePrompt = true;
			dialog.ValidateNames = true;
			dialog.Title = "Save Transcript As...";
			DialogResult result = dialog.ShowDialog(this);
			if (result == DialogResult.OK)
			{
				Cursor c = this.Cursor;
				this.Cursor = Cursors.WaitCursor;
				try
				{
					if (!Path.GetDirectoryName(dialog.FileName).Equals(Environment.GetFolderPath(Environment.SpecialFolder.Personal)))
					{
						_LastFolder = Path.GetDirectoryName(dialog.FileName);
					}
					switch (Path.GetExtension(dialog.FileName))
					{
						case ".stt":
							// write in our binary format
							WriteTranscriptToFile(OpenFile(dialog.FileName));
							break;
						case ".csv":
							// write as a CSV file
							WriteTranscriptToCsvFile(OpenFile(dialog.FileName));
							break;
						case ".htm":
						case ".html":
							// write as an HTML file
							WriteTranscriptToHtmlFile(OpenFile(dialog.FileName));
							break;
						default:
							// write as a straight text log
							WriteTranscriptToTextFile(OpenFile(dialog.FileName));
							break;
					}
				}
				catch
				{
					throw;
				}
				finally
				{
					this.Cursor = c;
				}
			}
		}

		private void FileMenu_OpenTranscript_Click(object sender, System.EventArgs e)
		{
			DialogResult result;
			if (_SerialPort.Connected)
			{
				result = MessageBox.Show(this, "You currently have a session connected.  Do you wish to save this session transcript first?", "Save current session?", MessageBoxButtons.YesNo);
				if (result == DialogResult.Yes)
				{
					FileMenu_SaveTranscript_Click(null, null);
				}
				DisconnectButton_Click(null, null);
			}
			OpenFileDialog dialog = new OpenFileDialog();
			dialog.CheckFileExists = true;
			dialog.CheckPathExists = true;
			dialog.DefaultExt = "stt";
			dialog.FileName = "SessionTranscript";
			dialog.Filter = "SharpTerminal Transcripts (*.stt)|*.stt|All Files (*.*)|*.*";
			dialog.FilterIndex = 1;
			if (_LastFolder.Length == 0)
			{
				dialog.InitialDirectory = _LastFolder;
			}
			else
			{
				dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			}
			dialog.ValidateNames = true;
			dialog.Title = "Open Transcript...";
			result = dialog.ShowDialog(this);
			if (result == DialogResult.OK)
			{
				Cursor c = this.Cursor;
				this.Cursor = Cursors.WaitCursor;
				StreamReader reader = null;
				try
				{
					if (!Path.GetDirectoryName(dialog.FileName).Equals(Environment.GetFolderPath(Environment.SpecialFolder.Personal)))
					{
						_LastFolder = Path.GetDirectoryName(dialog.FileName);
					}
					// Deserialize the file
					reader = new StreamReader(dialog.FileName, Encoding.UTF8);
					XmlSerializer s = new XmlSerializer(typeof(Message[]));
					Message[] messages = s.Deserialize(reader) as Message[];
					if (messages != null)
					{
						_Transcript.Clear();
						foreach (Message m in messages)
						{
							_Transcript.Add(m);
						}
						ReDisplay();
					}
					else
					{
						MessageBox.Show(this, "There were no entries in the transcript file.  If you expected entries, this may not be a valid SharpTerminal transcript file.");
					}
				}
				catch (Exception)
				{
					MessageBox.Show(this, "Unable to deserialize the transcript file.  It may not be a valid SharpTerminal transcript file.");
				}
				finally
				{
					try
					{
						if (reader != null) reader.Close();
					}
					catch
					{
					}
					this.Cursor = c;
				}
			}
		}

		private StreamWriter OpenFile(string filename)
		{
			return new StreamWriter(filename, false, Encoding.UTF8);
		}

		private void WriteTranscriptToFile(StreamWriter writer)
		{
			try
			{
				Message[] messages = new Message[_Transcript.Count];
				_Transcript.CopyTo(messages);
				XmlSerializer s = new XmlSerializer(typeof(Message[]));
				s.Serialize(writer, messages);
			}
			catch
			{
				throw;
			}
			finally
			{
				writer.Close();
			}
		}

		private void WriteTranscriptToCsvFile(StreamWriter writer)
		{
			try
			{
				writer.WriteLine("\"Timestamp\", \"Direction\", \"Data\"");
				foreach (Message m in _Transcript)
				{
					writer.WriteLine("\"{0}\", \"{1}\", \"{2}\"", m.Timestamp.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"), m.Direction.ToString(), Encoding.UTF8.GetString(m.Data));
				}
			}
			catch
			{
				throw;
			}
			finally
			{
				writer.Close();
			}
		}

		private void WriteTranscriptToHtmlFile(StreamWriter writer)
		{
			try
			{
				writer.Write("<HTML>{0}</HTML>", _Body.outerHTML);
			}
			catch
			{
				throw;
			}
			finally
			{
				writer.Close();
			}
		}

		private void WriteTranscriptToTextFile(StreamWriter writer)
		{
			try
			{
				foreach (Message m in _Transcript)
				{
					writer.Write(Encoding.UTF8.GetString(m.Data));
				}
			}
			catch
			{
				throw;
			}
			finally
			{
				writer.Close();
			}
		}

		private void HelpMenu_About_Click(object sender, System.EventArgs e)
		{
			STAboutForm about = new STAboutForm();
			about.ShowDialog(this);
		}

		private void SaveDefaultSettingsButton_Click(object sender, System.EventArgs e)
		{
			XmlDocument doc = new XmlDocument();
			XmlNode configNode = doc.CreateNode(XmlNodeType.Element, "SharpTerminal", "");
			doc.AppendChild(configNode);

			XmlAttribute attr = doc.CreateAttribute("", "Version", "");
			attr.Value = Application.ProductVersion;
			configNode.Attributes.Append(attr);

			XmlNode node = doc.CreateNode(XmlNodeType.Element, "ComPort", "");
			configNode.AppendChild(node);
			node.InnerText = ComPort.Text;

			node = doc.CreateNode(XmlNodeType.Element, "BaudRate", "");
			configNode.AppendChild(node);
			node.InnerText = BaudRateList.Text;

			node = doc.CreateNode(XmlNodeType.Element, "DataBits", "");
			configNode.AppendChild(node);
			node.InnerText = DataSizeList.Text;

			node = doc.CreateNode(XmlNodeType.Element, "Parity", "");
			configNode.AppendChild(node);
			node.InnerText = ParityList.Text;

			node = doc.CreateNode(XmlNodeType.Element, "StopBits", "");
			configNode.AppendChild(node);
			node.InnerText = StopBitsList.Text;

			node = doc.CreateNode(XmlNodeType.Element, "FlowControl", "");
			configNode.AppendChild(node);
			node.InnerText = FlowControlList.Text;

			node = doc.CreateNode(XmlNodeType.Element, "DisplayCR", "");
			configNode.AppendChild(node);
			node.InnerText = ShowCR.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "DisplayLF", "");
			configNode.AppendChild(node);
			node.InnerText = ShowLF.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "DisplayCRLF", "");
			configNode.AppendChild(node);
			node.InnerText = ShowCRLF.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "AutoAddLineEnd", "");
			configNode.AppendChild(node);
			node.InnerText = AutoAddLineEnd.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "SelectedLineEnd", "");
			configNode.AppendChild(node);
			node.InnerText = LineEndChars.Text;

			node = doc.CreateNode(XmlNodeType.Element, "ShowRTS", "");
			configNode.AppendChild(node);
			node.InnerText = ShowRtsTransitions.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "ShowCTS", "");
			configNode.AppendChild(node);
			node.InnerText = ShowCtsTransitions.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "ShowConnections", "");
			configNode.AppendChild(node);
			node.InnerText = ShowConnections.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "EchoLocal", "");
			configNode.AppendChild(node);
			node.InnerText = EchoLocal.Checked.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "DisplayMode", "");
			configNode.AppendChild(node);
			node.InnerText = _DisplayMode.ToString();

			node = doc.CreateNode(XmlNodeType.Element, "LastFolder", "");
			configNode.AppendChild(node);
			node.InnerText = _LastFolder;

			string filename = "";
			filename = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
															"SharpTerminal");
			if (!Directory.Exists(filename))
			{
				Directory.CreateDirectory(filename);
			}
			filename = Path.Combine(filename, "configuration.xml");

			doc.Save(filename);
		}

		private void LoadDefaultSettings()
		{
			string filename = "";
			filename = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
				"SharpTerminal");
			filename = Path.Combine(filename, "configuration.xml");

			if (File.Exists(filename))
			{
				XmlDocument doc = new XmlDocument();
				doc.Load(filename);
				XmlNode configNode = doc.SelectSingleNode("SharpTerminal");

				string version = configNode.Attributes["Version"].Value;
				string[] versionInfo = version.Split('.');
				string myVersion = Application.ProductVersion;
				string[] myVersionInfo = myVersion.Split('.');
				if (int.Parse(myVersionInfo[0]) > int.Parse(versionInfo[0]) ||
						int.Parse(myVersionInfo[1]) > int.Parse(versionInfo[1]))
				{
					ConvertConfiguration(versionInfo[0], versionInfo[1]);
				}

				BaudRateList.SelectedIndex = BaudRateList.FindStringExact(configNode.SelectSingleNode("BaudRate").InnerText);
				_SerialPort.BaudRate = (SerialPort.BaudRates)uint.Parse(BaudRateList.Text);

				DataSizeList.SelectedIndex = DataSizeList.FindStringExact(configNode.SelectSingleNode("DataBits").InnerText);
				_SerialPort.DataBits = (SerialPort.DataSizes)uint.Parse(DataSizeList.Text);

				ParityList.SelectedIndex = ParityList.FindStringExact(configNode.SelectSingleNode("Parity").InnerText);
				_SerialPort.Parity = (SerialPort.Parities)ParityList.SelectedIndex;

				StopBitsList.SelectedIndex = StopBitsList.FindStringExact(configNode.SelectSingleNode("StopBits").InnerText);
				_SerialPort.StopBits = (SerialPort.StopSizes)StopBitsList.SelectedIndex;

				FlowControlList.SelectedIndex = FlowControlList.FindStringExact(configNode.SelectSingleNode("FlowControl").InnerText);
				_SerialPort.FlowControl = (SerialPort.FlowControlTypes)FlowControlList.SelectedIndex;

				ShowCR.Checked = bool.Parse(configNode.SelectSingleNode("DisplayCR").InnerText);
				ShowLF.Checked = bool.Parse(configNode.SelectSingleNode("DisplayLF").InnerText);
				ShowCRLF.Checked = bool.Parse(configNode.SelectSingleNode("DisplayCRLF").InnerText);
				AutoAddLineEnd.Checked = (configNode.SelectSingleNode("AutoAddLineEnd") == null ? false : bool.Parse(configNode.SelectSingleNode("AutoAddLineEnd").InnerText));
				LineEndChars.Text = (configNode.SelectSingleNode("SelectedLineEnd") == null ? "" : configNode.SelectSingleNode("SelectedLineEnd").InnerText);
				ShowRtsTransitions.Checked = bool.Parse(configNode.SelectSingleNode("ShowRTS").InnerText);
				ShowCtsTransitions.Checked = bool.Parse(configNode.SelectSingleNode("ShowCTS").InnerText);
				ShowConnections.Checked = bool.Parse(configNode.SelectSingleNode("ShowConnections").InnerText);
				EchoLocal.Checked = (configNode.SelectSingleNode("EchoLocal") == null ? true : bool.Parse(configNode.SelectSingleNode("EchoLocal").InnerText));
				_DisplayMode = (configNode.SelectSingleNode("DisplayMode").InnerText.Equals("Text") ? Modes.Text : Modes.Hexadecimal);
				if (_DisplayMode == Modes.Text)
				{
					TextMode.Checked = true;
				}
				else
				{
					HexMode.Checked = true;
				}

				_LastFolder = configNode.SelectSingleNode("ShowConnections").InnerText;
			}
			else
			{
				BaudRateList.SelectedIndex = BaudRateList.FindStringExact("9600");
				DataSizeList.SelectedIndex = DataSizeList.FindStringExact("8");
				ParityList.SelectedIndex = ParityList.FindStringExact("None");
				StopBitsList.SelectedIndex = StopBitsList.FindStringExact("1");
				FlowControlList.SelectedIndex = FlowControlList.FindStringExact("None");

				_SerialPort.BaudRate = SerialPort.BaudRates.CBR_9600;
				_SerialPort.DataBits = SerialPort.DataSizes.Eight;
				_SerialPort.Parity = SerialPort.Parities.None;
				_SerialPort.StopBits = SerialPort.StopSizes.One;
				_SerialPort.FlowControl = SerialPort.FlowControlTypes.None;

				EchoLocal.Checked	= true;
				TextMode.Checked = true;
			}
		}

		private void ConvertConfiguration(string majorVersion, string minorVersion)
		{
			switch (int.Parse(majorVersion))
			{
				case 1:
					switch (int.Parse(minorVersion))
					{
						case 0:
							// Convert the original config file to a current one
							break;
						case 1:
							// Same version
							break;
					}
					break;
				default:
					MessageBox.Show(this, "Unable to figure out what version of SharpTerminal you used to have, cannot convert config file.");
					break;
			}
		}

		private void BaudRateList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			_CommSettingsChanged = true;
		}

		private void DataSizeList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			_CommSettingsChanged = true;
		}

		private void ParityList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			_CommSettingsChanged = true;
		}

		private void StopBitsList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			_CommSettingsChanged = true;
		}

		private void FlowControlList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			_CommSettingsChanged = true;
		}

		private void ClearButton_Click(object sender, System.EventArgs e)
		{
			DialogResult result = MessageBox.Show(this, "Do you wish to erase the session transcript as well?", "Erase transcript?", MessageBoxButtons.YesNo);
			if (result == DialogResult.Yes)
			{
				_Transcript.Clear();
			}
			_Body.innerHTML = "";
		}
		[DllImport("User32.dll")]
		static extern Boolean MessageBeep(uint beepType);
	} // class

	public enum Directions
	{
		Control,
		Read,
		Write
	}

	[XmlRoot("Message")]
	public class Message
	{
		[XmlElement("Data")]
		public byte[] Data;
		[XmlAttribute("Timestamp")]
		public DateTime Timestamp;
		[XmlAttribute("Direction")]
		public Directions Direction;
		public Message()
		{
		}
		public Message(DateTime timestamp, Directions direction, byte[] data)
		{
			this.Data = data;
			this.Timestamp = timestamp;
			this.Direction = direction;
		}
	} // class
} // namespace
