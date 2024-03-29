﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.IO;

namespace ComponenteWIA
{

	public partial class FormPrincipal : Form
	{

		int currentPage = 0;
		List<KeyValuePair<string, Image>> images;
		private string f_path = @"C:\temp\pictures\";
		private bool savedOrNot = false;
		private List<string> fNames;
		public List<string> fileNames
		{
			get { return fNames; }
			set { fNames = value; }
		}
		public String SavePath
		{
			get { return f_path; }
			set { f_path = value; }
		}
		public bool IsSaved
		{
			get { return savedOrNot; }
			set { savedOrNot = value; }
		}

		int cropX, cropY, cropWidth, cropHeight;
		//here rectangle border pen color=red and size=2;
		Pen borderpen = new Pen(Color.Red, 2);
		Bitmap currentImage;
		//fill the rectangle color =white
		SolidBrush rectbrush = new SolidBrush(Color.FromArgb(100, Color.White));
		int pages;


		private TextBox lblPath;
		private ListBox lbDevices;
		private IContainer components;
		private Button btnSave;
		private Button btnScanPreview;
		private Label label2;
		private Label label1;
		private ComboBox cFonte;
		private Label label3;
		private Label label4;
		private ComboBox cTamanho;
		private PictureBox pic_scan;

		public FormPrincipal()
		{
			InitializeComponent();
			IsSaved = false;
		}

		// populate the list of available scanners from WIA
		List<WIADeviceInfo> devices;

		protected override void OnLoad(EventArgs e)
		{
			lblPath.Text = SavePath;
			

			//get list of devices available
			devices = WIAScanner.ListaDispositivos().ToList();

			foreach (var device in devices)
			{
				//+ "(" + device.Server + ")"
				lbDevices.Items.Add(device.Name );
			}
			//check if device is not available
			if (lbDevices.Items.Count != 0)
			{
				lbDevices.SelectedIndex = 0;
			}

			cTamanho.Items.Add((WIAPageSize)WIAPageSize.A4);
			cTamanho.Items.Add((WIAPageSize)WIAPageSize.Legal);
			cTamanho.Items.Add((WIAPageSize)WIAPageSize.Letter);
			cTamanho.SelectedIndex = 0;
			/*
			cTamanho.Items.Add(new KeyValuePair<WIAPageSize, string>(WIAPageSize.A4, "A4"));
			cTamanho.Items.Add(new KeyValuePair<WIAPageSize, string>(WIAPageSize.Legal, "Legal"));
			cTamanho.Items.Add(new KeyValuePair<WIAPageSize, string>(WIAPageSize.Letter, "Letter"));
			*/
			cFonte.Items.Add((WIAScanner.TipoLeituraDocumento)WIAScanner.TipoLeituraDocumento.Feeder);
			cFonte.Items.Add((WIAScanner.TipoLeituraDocumento)WIAScanner.TipoLeituraDocumento.FeederDuplex);
			cFonte.Items.Add((WIAScanner.TipoLeituraDocumento)WIAScanner.TipoLeituraDocumento.FlatBed);
			cFonte.SelectedIndex = 0;
			base.OnLoad(e);
		}

		private void doScan(WIAScanQuality scanQuality)
		{
			try
			{
				//get list of devices available

				if (lbDevices.Items.Count == 0)
				{
					MessageBox.Show("You do not have any WIA devices.");

				}
				else
				{
					// get the selected scanner
					var device = devices[lbDevices.SelectedIndex];

					//get images from scanner
					//var pages_to_scan = 2;  
					WIAScanner.CaminhoArquivo = lblPath.Text ;// @"C:\temp\pictures\";
																										//images = WIAScanner.Scan(device.DeviceID, WIAPageSize.A4,WIAScanner.TipoLeituraDocumento.FeederDuplex);
					images = WIAScanner.Scan(device.DeviceID, (WIAPageSize)cTamanho.SelectedItem, (WIAScanner.TipoLeituraDocumento) cFonte.SelectedItem);
					pages = images.Count;
					HashSet<string> mensagens = WIAScanner.Mensagens;
					string ret = "";
					foreach (string mens in mensagens)
					{
						ret = ret + "\n" + mens;
					}
					if (images != null)
					{
						foreach (KeyValuePair<string, Image> image in images)
						{
							pic_scan.Image = image.Value;
							pic_scan.Show();
							pic_scan.SizeMode = PictureBoxSizeMode.StretchImage;
							currentImage = new Bitmap(image.Value);
							btnSave.Enabled = true;
							currentPage = 0;
						}
					}
					if(ret.Length > 0 ) MessageBox.Show(ret);
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}

		}

		private void btnScanPreview_Click(object sender, EventArgs e)
		{
			this.doScan(WIAScanQuality.Preview);
		}
		private void btnScanHighQuality_Click(object sender, EventArgs e)
		{
			this.doScan(WIAScanQuality.Final);
		}


		private string uniqueOutputFilepath()
		{
			string save_dir = lblPath.Text;
			string currentFName;
			string save_fullpath;
			int docno = 0;
			do
			{
				currentFName = String.Format("scan-{0}.jpeg", docno);
				save_fullpath = Path.Combine(save_dir, currentFName);
				docno++;
			} while (File.Exists(save_fullpath));
			return save_fullpath;
		}

		private void btnSave_Click(object sender, EventArgs e)
		{
			try
			{
				if (currentImage != null)
				{
					string save_fullpath = this.uniqueOutputFilepath();
					currentImage.Save(save_fullpath, ImageFormat.Jpeg);
					MessageBox.Show("Document Saved Successfully: " + save_fullpath);
					IsSaved = true;
					currentPage += 1;
					if (currentPage < (pages))
					{
						pic_scan.Image = images[currentPage].Value;
						pic_scan.Show();
						pic_scan.SizeMode = PictureBoxSizeMode.StretchImage;
						currentImage = new Bitmap(images[currentPage].Value);
						btnSave.Enabled = true;
					}
					else
					{
						btnSave.Enabled = false;
					}
				}
			}
			catch (Exception exc)
			{
				IsSaved = false;
				MessageBox.Show(exc.Message);
			}
		}


		private void InitializeComponent()
		{
			this.lblPath = new System.Windows.Forms.TextBox();
			this.lbDevices = new System.Windows.Forms.ListBox();
			this.btnSave = new System.Windows.Forms.Button();
			this.pic_scan = new System.Windows.Forms.PictureBox();
			this.btnScanPreview = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cFonte = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cTamanho = new System.Windows.Forms.ComboBox();
			((System.ComponentModel.ISupportInitialize)(this.pic_scan)).BeginInit();
			this.SuspendLayout();
			// 
			// lblPath
			// 
			this.lblPath.Location = new System.Drawing.Point(12, 327);
			this.lblPath.Name = "lblPath";
			this.lblPath.Size = new System.Drawing.Size(216, 20);
			this.lblPath.TabIndex = 0;
			// 
			// lbDevices
			// 
			this.lbDevices.FormattingEnabled = true;
			this.lbDevices.Location = new System.Drawing.Point(12, 76);
			this.lbDevices.Name = "lbDevices";
			this.lbDevices.Size = new System.Drawing.Size(256, 108);
			this.lbDevices.TabIndex = 3;
			// 
			// btnSave
			// 
			this.btnSave.Location = new System.Drawing.Point(12, 353);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(75, 23);
			this.btnSave.TabIndex = 5;
			this.btnSave.Text = "btnSave";
			this.btnSave.UseVisualStyleBackColor = true;
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// pic_scan
			// 
			this.pic_scan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pic_scan.Location = new System.Drawing.Point(283, 76);
			this.pic_scan.Name = "pic_scan";
			this.pic_scan.Size = new System.Drawing.Size(343, 408);
			this.pic_scan.TabIndex = 6;
			this.pic_scan.TabStop = false;
			// 
			// btnScanPreview
			// 
			this.btnScanPreview.Location = new System.Drawing.Point(12, 276);
			this.btnScanPreview.Name = "btnScanPreview";
			this.btnScanPreview.Size = new System.Drawing.Size(136, 23);
			this.btnScanPreview.TabIndex = 7;
			this.btnScanPreview.Text = "Scan";
			this.btnScanPreview.UseVisualStyleBackColor = true;
			this.btnScanPreview.Click += new System.EventHandler(this.btnScanPreview_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(104, 60);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(66, 13);
			this.label2.TabIndex = 8;
			this.label2.Text = "Scanner List";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(9, 311);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(57, 13);
			this.label1.TabIndex = 9;
			this.label1.Text = "SavePath:";
			// 
			// cFonte
			// 
			this.cFonte.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cFonte.FormattingEnabled = true;
			this.cFonte.Location = new System.Drawing.Point(15, 209);
			this.cFonte.Name = "cFonte";
			this.cFonte.Size = new System.Drawing.Size(253, 21);
			this.cFonte.TabIndex = 10;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 193);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 13);
			this.label3.TabIndex = 11;
			this.label3.Text = "Fonte";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(12, 233);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(97, 13);
			this.label4.TabIndex = 12;
			this.label4.Text = "Tamanho do Papel";
			// 
			// cTamanho
			// 
			this.cTamanho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cTamanho.FormattingEnabled = true;
			this.cTamanho.Location = new System.Drawing.Point(15, 250);
			this.cTamanho.Name = "cTamanho";
			this.cTamanho.Size = new System.Drawing.Size(253, 21);
			this.cTamanho.TabIndex = 13;
			// 
			// MainForm
			// 
			this.ClientSize = new System.Drawing.Size(670, 531);
			this.Controls.Add(this.cTamanho);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cFonte);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnScanPreview);
			this.Controls.Add(this.pic_scan);
			this.Controls.Add(this.btnSave);
			this.Controls.Add(this.lbDevices);
			this.Controls.Add(this.lblPath);
			this.Name = "MainForm";
			((System.ComponentModel.ISupportInitialize)(this.pic_scan)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

	}
}
