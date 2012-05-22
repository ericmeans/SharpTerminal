using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace SharpTerminal
{
	/// <summary>
	/// Summary description for STAboutForm.
	/// </summary>
	public class STAboutForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button OKButton;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Label label2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public STAboutForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
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
			this.label1 = new System.Windows.Forms.Label();
			this.OKButton = new System.Windows.Forms.Button();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Dock = System.Windows.Forms.DockStyle.Top;
			this.label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label1.Font = new System.Drawing.Font("Eric", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(0, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(402, 80);
			this.label1.TabIndex = 0;
			this.label1.Text = "SharpTerminal";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// OKButton
			// 
			this.OKButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.OKButton.Location = new System.Drawing.Point(164, 144);
			this.OKButton.Name = "OKButton";
			this.OKButton.TabIndex = 1;
			this.OKButton.Text = "OK";
			this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
			// 
			// linkLabel1
			// 
			this.linkLabel1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.linkLabel1.Location = new System.Drawing.Point(16, 112);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(376, 23);
			this.linkLabel1.TabIndex = 3;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "http://www.randomtree.org/sharpterminal/";
			this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// label2
			// 
			this.label2.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label2.Location = new System.Drawing.Point(105, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(192, 23);
			this.label2.TabIndex = 4;
			this.label2.Text = "Copyright (c) 2004 Eric Means";
			// 
			// STAboutForm
			// 
			this.AcceptButton = this.OKButton;
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(402, 178);
			this.ControlBox = false;
			this.Controls.Add(this.label2);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.OKButton);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "STAboutForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Load += new System.EventHandler(this.STAboutForm_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void OKButton_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			// Determine which link was clicked within the LinkLabel.
			this.linkLabel1.Links[linkLabel1.Links.IndexOf(e.Link)].Visited = true;

			// Display the appropriate link based on the value of the 
			// LinkData property of the Link object.
			string target = e.Link.LinkData as string;

			// If the value looks like a URL, navigate to it.
			// Otherwise, display it in a message box.
			if(null != target && target.StartsWith("http://"))
			{
				System.Diagnostics.Process.Start(target);
			}
		}

		private void STAboutForm_Load(object sender, System.EventArgs e)
		{
			this.linkLabel1.Links[0].LinkData = this.linkLabel1.Text;
		}
	}
}
