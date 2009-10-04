using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text;
using System.Diagnostics;

namespace PdfMC
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmMrgCpy : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnMerge;
		private System.Windows.Forms.TextBox txtMonth;
		private System.Windows.Forms.TextBox txtYear;
		private System.Windows.Forms.Label lblMonth;
		private System.Windows.Forms.Label lblYear;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.Label lblStatus;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMrgCpy()
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
			this.btnMerge = new System.Windows.Forms.Button();
			this.txtMonth = new System.Windows.Forms.TextBox();
			this.txtYear = new System.Windows.Forms.TextBox();
			this.lblMonth = new System.Windows.Forms.Label();
			this.lblYear = new System.Windows.Forms.Label();
			this.btnExit = new System.Windows.Forms.Button();
			this.lblStatus = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnMerge
			// 
			this.btnMerge.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.btnMerge.ForeColor = System.Drawing.Color.Navy;
			this.btnMerge.Location = new System.Drawing.Point(176, 16);
			this.btnMerge.Name = "btnMerge";
			this.btnMerge.Size = new System.Drawing.Size(136, 24);
			this.btnMerge.TabIndex = 3;
			this.btnMerge.Text = "Merge";
			this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
			// 
			// txtMonth
			// 
			this.txtMonth.Location = new System.Drawing.Point(72, 16);
			this.txtMonth.Name = "txtMonth";
			this.txtMonth.Size = new System.Drawing.Size(48, 20);
			this.txtMonth.TabIndex = 1;
			this.txtMonth.Text = "";
			// 
			// txtYear
			// 
			this.txtYear.Location = new System.Drawing.Point(72, 40);
			this.txtYear.Name = "txtYear";
			this.txtYear.Size = new System.Drawing.Size(48, 20);
			this.txtYear.TabIndex = 2;
			this.txtYear.Text = "";
			// 
			// lblMonth
			// 
			this.lblMonth.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.lblMonth.Location = new System.Drawing.Point(8, 16);
			this.lblMonth.Name = "lblMonth";
			this.lblMonth.Size = new System.Drawing.Size(40, 16);
			this.lblMonth.TabIndex = 3;
			this.lblMonth.Text = "Month";
			// 
			// lblYear
			// 
			this.lblYear.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.lblYear.Location = new System.Drawing.Point(8, 40);
			this.lblYear.Name = "lblYear";
			this.lblYear.Size = new System.Drawing.Size(40, 16);
			this.lblYear.TabIndex = 4;
			this.lblYear.Text = "Year";
			// 
			// btnExit
			// 
			this.btnExit.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.btnExit.ForeColor = System.Drawing.Color.Navy;
			this.btnExit.Location = new System.Drawing.Point(176, 40);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(136, 23);
			this.btnExit.TabIndex = 4;
			this.btnExit.Text = "Exit";
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// lblStatus
			// 
			this.lblStatus.ForeColor = System.Drawing.Color.Maroon;
			this.lblStatus.Location = new System.Drawing.Point(16, 72);
			this.lblStatus.Name = "lblStatus";
			this.lblStatus.Size = new System.Drawing.Size(296, 24);
			this.lblStatus.TabIndex = 5;
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Tahoma", 8F);
			this.label1.Location = new System.Drawing.Point(128, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 6;
			this.label1.Text = "(01-12)";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Tahoma", 8F);
			this.label2.Location = new System.Drawing.Point(128, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 7;
			this.label2.Text = "(YYYY)";
			// 
			// frmMrgCpy
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(328, 110);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblStatus);
			this.Controls.Add(this.btnExit);
			this.Controls.Add(this.lblYear);
			this.Controls.Add(this.lblMonth);
			this.Controls.Add(this.txtYear);
			this.Controls.Add(this.txtMonth);
			this.Controls.Add(this.btnMerge);
			this.Name = "frmMrgCpy";
			this.Text = "PPL Merge/Copy";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMrgCpy());
		}

		private void btnMerge_Click(object sender, System.EventArgs e)
		{
			this.lblStatus.Text = "Copying Files";
			this.lblStatus.Refresh();

			int month = Int32.MinValue;
			int year = Int32.MinValue;

			string cMonth = this.txtMonth.Text;
			string cYear = this.txtYear.Text;

			try
			{
				month = Int32.Parse(this.txtMonth.Text);
				year = Int32.Parse(this.txtYear.Text);
			}
			catch(Exception)
			{
				this.lblStatus.Text = "Year &/or Month not valid integer";
				MessageBox.Show("Year and/or Month is not a valid integer", "PPL Merge/Copy", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (month < 1 || month > 12)
			{
				this.lblStatus.Text = "Not a valid month";
				MessageBox.Show("The month is not valid", "PPL Merge/Copy", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Local location of files to be merged
			string fileLocSrce = @"C:\ppls_tmp\";
			this.lblStatus.Text = "Copying Files";

			try
			{			
				//CopyOriginals(cMonth,cYear);

				// Create local temp directory and copy source files
				CreateTempDir(fileLocSrce, cMonth, cYear);

				// Collection to hold filenames to be merged for same facility
				ArrayList sameFacilityFiles = new ArrayList(6);

				string lastBusUnit = null;
				string busUnit = null;
				string facName = null;
				string lastFacName = null;

				// Get a directory listing of PDF files
				string[] dirs = Directory.GetFiles(fileLocSrce,"*.pdf");

				// Cycle through the files in the dirs array
				// fullFileName contains the entire path
				foreach (string fullFileName in dirs)
				{
					// Get the file name (without path) so the business unit can be extracted
					string fileName = Path.GetFileName(fullFileName);
					int usIndex = fileName.IndexOf('_');

					busUnit = fileName.Substring(0,5);
					facName = fileName.Substring(6, usIndex - 6);
				
					// If the last bu is null set it to current bu
					if (lastBusUnit == null)
					{
						lastBusUnit = busUnit; 
						lastFacName = facName;
					}

					// If current bu is not the same as the last bu, merge the files collected
					// for the current bu and clear the array list so it can start over with 
					// the next facility files
					if (busUnit != lastBusUnit)
					{
						MergeFiles(sameFacilityFiles, lastBusUnit, lastFacName, cMonth, cYear);
						this.lblStatus.Text = "Merged files for " + lastFacName;
						sameFacilityFiles.Clear();
						lastBusUnit = null;
					}

					// If the current bu is the same as the last bu, add to the arraylist
					sameFacilityFiles.Add(fullFileName);	

				}
				// Merges the last facility files
				MergeFiles(sameFacilityFiles, lastBusUnit, lastFacName, cMonth, cYear);
				this.lblStatus.Text = "Process is Complete";
			}
			finally
			{
				if (Directory.Exists(fileLocSrce))
					Directory.Delete(fileLocSrce, true);
			}
		}

		void MergeFiles(ArrayList mrgFiles, string bu, string fac, string mo, string yr)
		{
			if (mrgFiles.Count <= 0)
			{
				this.lblStatus.Text = "No files to Merge";
				MessageBox.Show ("There are no files to merge", "PPL Merge/Copy", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Location of the file that will contain the list of files to merge
			// The file is used as an argument to the PDFSplitMerge utility
			string mFile = @"c:\temp\FileList.txt";
            
			// Creates the file initialized above
			StreamWriter SW;
			File.Delete(mFile);
			SW=File.CreateText(mFile);

			// This loop basically writes the actual file names to the FileList.Txt file
			foreach (string file2Merge in mrgFiles)
			{
				SW.WriteLine(file2Merge);
			}
			SW.Close();


			// Sets the target file name format to be used in the arguments of the 
			// PDFSplitMerge utility
			string mtsMonth = yr + "-" + mo + "\\";

//			string vistaFilename = @"I:\OSC\Vista\PP_Coll\" + mtsMonth + "PPL_" + bu + ".pdf";
			string vistaFilename = @"L:\PP_Coll\" + mtsMonth + "PPL_" + bu + ".pdf";
			string storeFilename = @"I:\OSC\Private Collection Letters\" + mtsMonth + fac + " Collection Letters.pdf";

			
			string procArgs = "FileList.txt " + vistaFilename;

			// Runs the PDFSplitMerge utility to merge the files
			Process proc = new System.Diagnostics.Process();
			proc.EnableRaisingEvents=false;
			proc.StartInfo.Arguments = procArgs;
			proc.StartInfo.WorkingDirectory = "c:\\temp\\";
			proc.StartInfo.FileName="pdfpg.exe";
			proc.Start();
			proc.WaitForExit();

			// Also copy the file to the storage location
			File.Copy(vistaFilename, storeFilename);
		}
	
		// Create local temp directory and copy source files
		void CreateTempDir(string locSrce, string mo, string yr)
		{
			// Location of files to be merged
			string fileNetSrce = @"I:\OSC\Private Collection Letters\MoveToStorage";
//Test		string fileNetSrce = @"C:\Test\Private Collection Letters\MoveToStorage";
			
			// Location to copy original files
			string cMtsMonth = yr + "-" + mo + "\\";
			string origTarget = @"I:\OSC\Private Collection Letters\Originals\" + cMtsMonth ;
//Test		string origTarget = @"C:\Test\Private Collection Letters\Originals\" + cMtsMonth ;

			string netFile;
			string locFile;
			string origFile;
	
			try
			{
				// Determine whether the local temp directory exists
				if (Directory.Exists(locSrce))
				{
					Directory.Delete(locSrce, true);
				}

				// Try to create the directory
				DirectoryInfo locdi = Directory.CreateDirectory(locSrce);


//Begin Testing -- creating directories for target files
				// Set variables to hold the directory location for the following:
				// 1.  Original files (before merge::1,1A,2,2A, etc)
				// 2.  Merged files for storage
				// 3.  Vista files
				string cOrigDir = @"I:\OSC\Private Collection Letters\Originals\" + cMtsMonth ;
				string cMtsDir = @"I:\OSC\Private Collection Letters\" + cMtsMonth ;
//				string cVistaDir = @"I:\OSC\Vista\PP_Coll\" + cMtsMonth ;
				string cVistaDir = @"L:\PP_Coll\" + cMtsMonth ;
								
				// Determine whether the directories exist
				// If they do,  stop the process, it may be that this process has been 
				// executed already or the incorrect month was input
				
				// Original files
				if (Directory.Exists(cOrigDir))
				{
					this.lblStatus.Text = "Directory already exists: " + "\n" + cOrigDir ;
					MessageBox.Show ("Directory already exists: " + "\n" + cOrigDir, "PPL Merge/Copy", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);

					this.txtMonth.Text = "";
					this.txtYear.Text = "";

					return;
				}

				// Storage files
				if (Directory.Exists(cMtsDir))
				{
					this.lblStatus.Text = "Directory already exists: " + "\n" + cMtsDir;
					MessageBox.Show ("Directory already exists: " + "\n" + cMtsDir, "PPL Merge/Copy", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Vista files
				if (Directory.Exists(cVistaDir))
				{
					this.lblStatus.Text = "Directory already exists: " + "\n" + cVistaDir;
					MessageBox.Show ("Directory already exists: " + "\n" + cVistaDir, "PPL Merge/Copy", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}


				try
				{
					// Create the needed directories
					DirectoryInfo orgdi = Directory.CreateDirectory(cOrigDir);
					DirectoryInfo mtsdi = Directory.CreateDirectory(cMtsDir);
					DirectoryInfo vstdi = Directory.CreateDirectory(cVistaDir);
				}
				catch
				{
					this.lblStatus.Text = "Could not create a target directory";
					MessageBox.Show("Could not create a target directory", "PPL Merge/Copy", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

//End Testing -- creating directories for target files


				// Copy the files from MoveToStorage to local drive and Originals\YYYY-MM
				string[] origFiles = Directory.GetFiles(fileNetSrce,"*.pdf");
				foreach (string origF in origFiles)
				{
					netFile = origF;
					locFile = locSrce + Path.GetFileName(netFile);
					origFile = origTarget + Path.GetFileName(netFile);

					File.Copy(netFile, locFile);
					File.Copy(netFile, origFile);
					
					File.Delete(netFile);
				}
			}
			catch
			{
				MessageBox.Show ("The process failed", "PPL Merge/Copy", 
				MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

	}
}
