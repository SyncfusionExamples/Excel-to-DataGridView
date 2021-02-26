using Syncfusion.XlsIO;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ExceltoDataGrid
{
	public partial class Form1 : Form
	{
		#region Constants
		private const string DEFAULTPATH = @"..\..\..\..\..\..\..\Common\Data\XlsIO\{0}";
		#endregion

		#region Fields
		private DataGridView dataGridView;
		private Button btnImport;
		private Label label;
		#endregion

		#region Initialize
		public Form1()
		{
			InitializeComponent();
		}
		#endregion

		#region Export Data from Excel to DataGrid
		private void btnImport_Click(object sender, System.EventArgs e)
		{
			//Initialize the Excel Engine
			using (ExcelEngine excelEngine = new ExcelEngine())
			{
				//Initialize Application
				IApplication application = excelEngine.Excel;

				//Set default version for application.
				application.DefaultVersion = ExcelVersion.Excel2013;

				//Open existing workbook with data entered
				Assembly assembly = typeof(Form1).GetTypeInfo().Assembly;
				Stream fileStream = assembly.GetManifestResourceStream("ExceltoDataGrid.Sample.xlsx");
				IWorkbook workbook = application.Workbooks.Open(fileStream);

				//Accessing first worksheet in the workbook
				IWorksheet worksheet = workbook.Worksheets[0];

				//Export data from Excel worksheet to DataTable
				DataTable customersTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

				//Load the data to DataGridView
				dataGridView.DataSource = customersTable;

				//No exception will be thrown if there are unsaved workbooks.
				excelEngine.ThrowNotSavedOnDestroy = false;
			}
		}
		#endregion

		#region HelperMethods
		/// <summary>
		/// Get the file path of input file and return the same
		/// </summary>
		/// <param name="inputPath">Input file</param>
		/// <returns>path of the input file</returns>
		private string GetTemplatePath(string inputFile)
		{
			return string.Format(DEFAULTPATH, inputFile);
		}
		#endregion

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnImport = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.label = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(485, 373);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(158, 36);
            this.btnImport.TabIndex = 2;
            this.btnImport.Text = "Excel to DataGridView";
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.Location = new System.Drawing.Point(15, 60);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(628, 292);
            this.dataGridView.TabIndex = 0;
            // 
            // label
            // 
            this.label.Location = new System.Drawing.Point(12, 9);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(412, 48);
            this.label.TabIndex = 1;
            this.label.Text = "Click the button to load Excel spreadsheet to DataGrid through Essential XlsIO.";
            this.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(662, 425);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.label);
            this.Controls.Add(this.btnImport);
            this.Name = "Form1";
            this.Text = "Excel to DataGridView";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			//SyncfusionLicenseProvider.RegisterLicense(DemoCommon.FindLicenseKey());
			Application.Run(new Form1());
		}
	}
}
