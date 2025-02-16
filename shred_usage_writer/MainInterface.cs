using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace shred_usage_writer
{
    public partial class MainInterface : Form
    {
        public XLWorkbook wb;
        public IXLWorksheet ws;
        public string solutionDirectory;
        public string ogWorkbook;
        int thisYear = DateTime.Now.Year;
        int thisMonth = DateTime.Now.Month;
        int thisDay = DateTime.Now.Day;
        public MainInterface()
        {
            //this.solutionDirectory = GetSolutionDirectoryInfo().ToString().Remove(GetSolutionDirectoryInfo().ToString().Length - 18);
            //Trace.WriteLine(solutionDirectory);
            //string yearDirectory = solutionDirectory + thisYear.ToString();
            //if (!Directory.Exists(yearDirectory))
            //{
            //    Directory.CreateDirectory(yearDirectory);
            //}
            //string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(thisMonth);
            //string monthDirectory = System.IO.Path.Combine(yearDirectory, monthName);
            //if (!Directory.Exists(monthDirectory))
            //{
            //    Directory.CreateDirectory(monthDirectory);
            //}
            //string filePath = System.IO.Path.Combine(monthDirectory, thisDay.ToString() + "-" + monthName + "_Shred_Usage_Output" + ".xlsx");
            //ogWorkbook = System.IO.Path.Combine(solutionDirectory, "blank.xlsx");
            //if (!File.Exists(filePath))
            //{
            //    Trace.WriteLine(ogWorkbook);
            //    File.Copy(ogWorkbook, filePath);

            //}
            //this.wb = new XLWorkbook(filePath);

            this.solutionDirectory = GetSolutionDirectoryInfo().ToString().Remove(GetSolutionDirectoryInfo().ToString().Length - 18);
            Trace.WriteLine(solutionDirectory);

            string yearDirectory = System.IO.Path.Combine(solutionDirectory, thisYear.ToString());
            if (!Directory.Exists(yearDirectory))
            {
                Directory.CreateDirectory(yearDirectory);
            }

            string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(thisMonth);
            string monthDirectory = System.IO.Path.Combine(yearDirectory, monthName);
            if (!Directory.Exists(monthDirectory))
            {
                Directory.CreateDirectory(monthDirectory);
            }

            string filePath = System.IO.Path.Combine(monthDirectory, $"{thisDay}-{monthName}_Shred_Usage_Output.xlsx");

            // Extract the embedded resource and get its path
            ogWorkbook = ExtractBlankExcelTemplate();

            if (!File.Exists(filePath))
            {
                File.Copy(ogWorkbook, filePath);
            }

            // Load the new workbook
            this.wb = new XLWorkbook(filePath);

            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    Trace.WriteLine(sheet.Name);
                }
            }



            InitializeComponent();
            this.Text = "Miceli Dairy Products - Shred Usage Reporting Tool";
            this.ShowIcon = false;
            InitializeComboBox();
            InitializeRightPanel();
            //int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            //this.Width = screenWidth;
            //this.Height = Screen.PrimaryScreen.WorkingArea.Height;
            this.WindowState = FormWindowState.Maximized; // Maximizes the window
            this.FormBorderStyle = FormBorderStyle.None;  // Hides the title bar
            errorProvider = new ErrorProvider();
        }

        private string ExtractBlankExcelTemplate()
        {
            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "BLANK1.xlsx");

            // Prevent unnecessary extraction if the file already exists
            if (!File.Exists(tempPath))
            {
                string resourceName = "shred_usage_writer.Resources.BLANK1.xlsx"; // Adjust with your namespace

                using (Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (FileStream fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                    else
                    {
                        throw new FileNotFoundException($"Embedded resource not found: {resourceName}");
                    }
                }
            }

            return tempPath;
        }


        FlowLayoutPanel rightPanel = new FlowLayoutPanel();
        TableLayoutPanel tableLayout;


        // Declare List
        private List<string> L = new List<string>();

        // Declare Dynamic Item Name
        private string itemString;

        // Declare Error Provider
        private ErrorProvider errorProvider;

        // Declare Controls
        internal MessageBox SubmitCheckBox;
        internal Button SubmitCheckBoxYes;
        internal Button SubmitCheckBoxNo;
        internal ComboBox ComboBox1;
        internal DateTimePicker Date;
        internal NumericUpDown ToteSkidNumber;
        internal NumericUpDown NumberPieces;
        internal NumericUpDown BinWeight;
        internal DateTimePicker StartTime;
        internal NumericUpDown Temp;
        internal Button SubmitButton;
        internal CheckBox BinSealGrade;
        internal GroupBox FirmnessBox;
        internal RadioButton FirmnessFirm;
        internal RadioButton FirmnessSoft;
        internal GroupBox DelvicidBox;
        internal RadioButton DelvicidTrue;
        internal RadioButton DelvicidFalse;
        internal TextBox Initials;
        internal NumericUpDown BagCount;
        internal TextBox PowderLotNumber;


        //  Declare Labels
        internal Label comboBoxLabel;
        internal Label dateLabel;
        internal Label skidNumberLabel;
        internal Label piecesNumberLabel;
        internal Label binWeightLabel;
        internal Label startTimeLabel;
        internal Label tempLabel;
        internal Label binSealLabel;
        internal Label firmnessLabel;
        internal Label delvicidLabel;
        internal Label initialsLabel;
        internal Label bagCountLabel;
        internal Label powderLotNumberLabel;

        public static DirectoryInfo? GetSolutionDirectoryInfo(string? currentPath = null)
        {
            DirectoryInfo? directory = new(
                currentPath ?? Directory.GetCurrentDirectory());
            while (directory != null && !directory.GetFiles("*.sln").Any())
            {
                directory = directory.Parent;
            }
            return directory;
        }

        private void InitializeComboBox()
        {
            this.ComboBox1 = new ComboBox();
            this.ComboBox1.Location = new System.Drawing.Point((this.ClientSize.Width / 5) * 2, 90);
            this.ComboBox1.Name = "ComboBox1";
            this.ComboBox1.Size = new System.Drawing.Size(360, 50);
            ComboBox1.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.ComboBox1.TabIndex = 0;
            this.ComboBox1.Text = "Select Block Item";
            string[] installs = new string[] { "001-000133", "001-000169", "001-000195", "001-000229", "001-000360", "001-000455", "001-000470", "001-000705",
            "001-000525", "001-000528", "008-000005 PS Purchased", "008-000021 WM Purchased", "008-000001 Asiago", "008-000002 Cheddar", "008-000006 Parmesan",
            "008-000010 White Ched", "008-000022 Meunster", "008-000007 Provolone", "002-000035 Scrap", "Powder"};
            ComboBox1.Items.AddRange(installs);
            ComboBox1.CausesValidation = false;
            this.Controls.Add(this.ComboBox1);

            comboBoxLabel = new Label();
            comboBoxLabel.Location = new System.Drawing.Point(((this.ClientSize.Width / 5) * 2) - 220, 90);
            comboBoxLabel.Size = new System.Drawing.Size(220, 50);
            comboBoxLabel.Name = "comboBoxLabel";
            comboBoxLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            comboBoxLabel.Text = "ITEM SELECT";
            this.Controls.Add(comboBoxLabel);



            // Hook up the event handler.
            this.ComboBox1.SelectedIndexChanged +=
                new System.EventHandler(ComboBox1_SelectedIndexChanged);
        }

        private void InitializeRightPanel()
        {
            rightPanel.Size = new Size((this.ClientSize.Width / 3) * 2, this.ClientSize.Height / 3); // Right half, top third
            rightPanel.Location = new System.Drawing.Point((this.ClientSize.Width / 5) * 4, 70); // Position at top-right
            rightPanel.BackColor = System.Drawing.Color.White;
            rightPanel.FlowDirection = FlowDirection.TopDown;
            rightPanel.BorderStyle = BorderStyle.FixedSingle;
            rightPanel.AutoScroll = true;

            this.Controls.Add(rightPanel);

            // Example list of strings
            List<string> items = new List<string> { "", "", "", "", "", "", "", "", "", "" };

            foreach (string item in items)
            {
                Label label = new Label();
                label.Text = item;
                label.ForeColor = System.Drawing.Color.Black;
                label.BackColor = System.Drawing.Color.White;
                label.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold); // Bigger text
                label.AutoSize = true;
                rightPanel.Controls.Add(label);
            }

            Label labelVersion = new Label();
            labelVersion.Location = new System.Drawing.Point((this.ClientSize.Width / 5) * 4, 50);
            labelVersion.Font = new System.Drawing.Font("Arial", 8, FontStyle.Regular);
            labelVersion.Size = new Size(500, 60);
            labelVersion.Text = "v1.0.0                                         PBS2025";
            labelVersion.ForeColor = System.Drawing.Color.Gray;
            this.Controls.Add(labelVersion);

        }


        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            string selectedProduct = (string)comboBox.SelectedItem;

            switch (selectedProduct)
            {
                case "001-000133":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000133      ";
                    InitializeBlockTypeA("1-133");
                    break;
                case "001-000169":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000169 Part-Skim Block";
                    InitializeBlockTypeA("1-169");
                    break;
                case "001-000195":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000195      ";
                    InitializeBlockTypeA("1-195");
                    break;
                case "001-000229":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000229      ";
                    InitializeBlockTypeA("1-229");
                    break;
                case "001-000360":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000360      ";
                    InitializeBlockTypeA("1-360");
                    break;
                case "001-000455":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000455      ";
                    InitializeBlockTypeA("1-455");
                    break;
                case "001-000470":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000470      ";
                    InitializeBlockTypeA("1-470");
                    break;
                case "001-000705":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000705      ";
                    InitializeBlockTypeA("1-705");
                    break;
                case "001-000525":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000525 Whole Milk Block";
                    InitializeBlockTypeA("1-525");
                    break;
                case "001-000528":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000528      ";
                    InitializeBlockTypeA("1-528");
                    break;
                case "008-000005 PS Purchased":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000005      Purchased PS Block";
                    InitializeBlockTypeA("PS Purchased");
                    break;
                case "008-000021 WM Purchased":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000021      Purchased WM Block";
                    InitializeBlockTypeA("WM Purchased");
                    break;
                case "008-000001 Asiago":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000001      ";
                    InitializeBlockTypeB("Asiago40#");
                    break;
                case "008-000002 Cheddar":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000002      ";
                    InitializeBlockTypeB("Ched40#");
                    break;
                case "008-000006 Parmesan":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000006      ";
                    InitializeBlockTypeB("Parm40#");
                    break;
                case "008-000010 White Ched":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000010      ";
                    InitializeBlockTypeB("WhiteChed40#");
                    break;
                case "008-000022 Meunster":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000022      ";
                    InitializeBlockTypeC("MuensterCS");
                    break;
                case "008-000007 Provolone":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     008-000007      ";
                    InitializeBlockTypeC("ProvLogCS");
                    break;
                case "002-000035 Scrap":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     002-000035      ";
                    InitializeScrap();
                    break;
                case "Powder":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     Powder      ";
                    InitializePowder();
                    break;
                default:
                    MessageBox.Show("Please Make a Product Selection");
                    break;

            }
        }

        private void NewSelection()
        {
            this.Controls.Remove(SubmitButton);
            this.Controls.Remove(initialsLabel);
            this.Controls.Remove(Initials);
            if (Initials != null) { Initials.Dispose(); }
            this.Controls.Remove(bagCountLabel);
            this.Controls.Remove(BagCount);
            if (BagCount != null) { BagCount.Dispose(); }
            this.Controls.Remove(powderLotNumberLabel);
            this.Controls.Remove(PowderLotNumber);
            if (PowderLotNumber != null) { PowderLotNumber.Dispose(); }
            this.Controls.Remove(tableLayout);
            if (tableLayout != null) { tableLayout.Dispose(); }
            errorProvider = new ErrorProvider();
        }

        private void InitializeBlockTypeA(string productNumber)
        {

            // Define a fixed space between rows (e.g., 10% for spacing)
            float spacePercentage = 10F;

            tableLayout = new TableLayoutPanel();
            tableLayout.ColumnCount = 2;
            tableLayout.RowCount = 9;
            tableLayout.Dock = DockStyle.None;  // Remove automatic docking
            tableLayout.AutoSize = true;
            tableLayout.Location = new System.Drawing.Point(300, 250); // Move it right (X=300) and down (Y=200)
            tableLayout.Width = this.Width / 2; // Take up about half the width
            tableLayout.Padding = new Padding(20);
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // Labels
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Controls
            tableLayout.RowStyles.Clear(); // Clear any default row styles

            for (int i = 0; i < tableLayout.RowCount; i++)
            {
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F)); // Adds space between rows

            }


            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            dateLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(220, 50);
            Date.Text = DateTime.Today.ToString("MM/dd/yyyy");
            this.Date.Validating += new CancelEventHandler(LotDate_Validating_Handler);

            skidNumberLabel = new Label();
            skidNumberLabel.Text = "Skid/Tote Number:";
            skidNumberLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            skidNumberLabel.Size = new System.Drawing.Size(350, 50);
            this.Controls.Add(skidNumberLabel);
            ToteSkidNumber = new NumericUpDown();
            ToteSkidNumber.Name = "Skid Number";
            ToteSkidNumber.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            ToteSkidNumber.Size = new System.Drawing.Size(90, 50);
            ToteSkidNumber.Maximum = 199;
            ToteSkidNumber.Minimum = 0;
            ToteSkidNumber.Value = 0;
            ToteSkidNumber.Text = "";
            ToteSkidNumber.Validating += ToteSkidNumber_Validating;

            this.piecesNumberLabel = new Label();
            piecesNumberLabel.Text = "Number of Pieces:";
            piecesNumberLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            piecesNumberLabel.Size = new System.Drawing.Size(350, 50);
            this.Controls.Add(piecesNumberLabel);
            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Name = "Number of Pieces";
            NumberPieces.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.NumberPieces.Minimum = 1;
            this.NumberPieces.Maximum = 200;
            this.NumberPieces.Value = 160;
            this.NumberPieces.Size = new System.Drawing.Size(90, 50);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Bin Weight (lbs.):";
            binWeightLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binWeightLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Name = "Bin Weight";
            BinWeight.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(150, 50);
            BinWeight.Validating += BinWeight_Validating;

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            startTimeLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            StartTime.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(120, 50);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F°):";
            tempLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            tempLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Name = "Temperature";
            Temp.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 80.00M;
            this.Temp.Value = 0.00M;
            this.Temp.Size = new System.Drawing.Size(90, 50);
            Temp.Text = "";
            Temp.Validating += Temp_Validating;

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Size = new System.Drawing.Size(180, 80);
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = ContentAlignment.MiddleCenter;
            SubmitButton.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeA(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "Verified mold is not present:";
            binSealLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binSealLabel.Size = new System.Drawing.Size(400, 100);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            BinSealGrade.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinSealGrade.Size = new System.Drawing.Size(30, 30);
            BinSealGrade.Validating += BinSealGrade_Validating;

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.Size = new System.Drawing.Size(200, 50);
            firmnessLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Size = new System.Drawing.Size(220, 70);
            this.FirmnessBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label firmLabel = new Label();
            firmLabel.Text = "Firm";
            firmLabel.Location = new System.Drawing.Point(20, 25);
            firmLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(firmLabel);
            this.FirmnessFirm = new RadioButton();
            this.FirmnessFirm.Name = "Firmness Firm";
            this.FirmnessFirm.Location = new System.Drawing.Point(80, 25);
            this.FirmnessFirm.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessFirm);
            this.FirmnessFirm.Checked = true;
            //  Second Button
            Label softLabel = new Label();
            softLabel.Text = "Soft";
            softLabel.Location = new System.Drawing.Point(110, 25);
            softLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(softLabel);
            this.FirmnessSoft = new RadioButton();
            this.FirmnessSoft.Name = "Firmness Soft";
            this.FirmnessSoft.Location = new System.Drawing.Point(170, 25);
            this.FirmnessSoft.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessSoft);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.Size = new System.Drawing.Size(300, 50);
            delvicidLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Size = new System.Drawing.Size(220, 70);
            this.DelvicidBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label trueDelvicidLabel = new Label();
            trueDelvicidLabel.Text = "Yes";
            trueDelvicidLabel.Location = new System.Drawing.Point(20, 25);
            trueDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(trueDelvicidLabel);
            this.DelvicidTrue = new RadioButton();
            this.DelvicidTrue.Name = "Delvicid True";
            this.DelvicidTrue.Location = new System.Drawing.Point(80, 25);
            this.DelvicidTrue.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidTrue);
            this.DelvicidTrue.Checked = true;
            //  Second Button
            Label falseDelvicidLabel = new Label();
            falseDelvicidLabel.Text = "No";
            falseDelvicidLabel.Location = new System.Drawing.Point(120, 25);
            falseDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(falseDelvicidLabel);
            this.DelvicidFalse = new RadioButton();
            this.DelvicidFalse.Name = "Delvicid False";
            this.DelvicidFalse.Location = new System.Drawing.Point(170, 25);
            this.DelvicidFalse.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidFalse);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Size = new System.Drawing.Size(120, 50);
            initialsLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            Initials.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Initials.Size = new System.Drawing.Size(70, 50);
            Initials.CharacterCasing = CharacterCasing.Upper;
            Initials.Validating += Initials_Validating;

            // Add Controls to TableLayoutPanel
            tableLayout.Controls.Add(dateLabel, 0, 0);
            tableLayout.Controls.Add(Date, 1, 0);
            tableLayout.Controls.Add(skidNumberLabel, 0, 1);
            tableLayout.Controls.Add(ToteSkidNumber, 1, 1);
            tableLayout.Controls.Add(piecesNumberLabel, 0, 2);
            tableLayout.Controls.Add(NumberPieces, 1, 2);
            tableLayout.Controls.Add(binWeightLabel, 0, 3);
            tableLayout.Controls.Add(BinWeight, 1, 3);
            tableLayout.Controls.Add(startTimeLabel, 0, 4);
            tableLayout.Controls.Add(StartTime, 1, 4);
            tableLayout.Controls.Add(tempLabel, 0, 5);
            tableLayout.Controls.Add(Temp, 1, 5);
            tableLayout.Controls.Add(firmnessLabel, 0, 6);
            tableLayout.Controls.Add(FirmnessBox, 1, 6);
            tableLayout.Controls.Add(delvicidLabel, 0, 7);
            tableLayout.Controls.Add(DelvicidBox, 1, 7);
            tableLayout.Controls.Add(initialsLabel, 0, 8);
            tableLayout.Controls.Add(Initials, 1, 8);
            tableLayout.Controls.Add(SubmitButton, 1, 10);
            tableLayout.Controls.Add(binSealLabel, 0, 9);
            tableLayout.Controls.Add(BinSealGrade, 1, 9);

            // Add TableLayoutPanel to Form
            this.Controls.Add(tableLayout);
        }

        private void InitializeBlockTypeB(string productNumber)
        {
            tableLayout = new TableLayoutPanel();
            tableLayout.ColumnCount = 2;
            tableLayout.RowCount = 8;
            tableLayout.Dock = DockStyle.None;  // Remove automatic docking
            tableLayout.AutoSize = true;
            tableLayout.Location = new System.Drawing.Point(300, 250); // Move it right (X=300) and down (Y=200)
            tableLayout.Width = this.Width / 2; // Take up about half the width
            tableLayout.Padding = new Padding(20);
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // Labels
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Controls
            tableLayout.RowStyles.Clear(); // Clear any default row styles

            for (int i = 0; i < tableLayout.RowCount; i++)
            {
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
            }

            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            dateLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(220, 50);
            Date.Text = DateTime.Today.ToString("MM/dd/yyyy");
            this.Date.Validating += new CancelEventHandler(LotDate_Validating_Handler);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Bin Weight (lbs.):";
            binWeightLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binWeightLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Name = "Bin Weight";
            BinWeight.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(150, 50);
            BinWeight.Validating += BinWeight_Validating;

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            startTimeLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            StartTime.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(120, 50);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F°):";
            tempLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            tempLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Name = "Temperature";
            Temp.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 80.00M;
            this.Temp.Value = 0.00M;
            this.Temp.Size = new System.Drawing.Size(90, 50);
            Temp.Text = "";
            Temp.Validating += Temp_Validating;

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Size = new System.Drawing.Size(180, 80);
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = ContentAlignment.MiddleCenter;
            SubmitButton.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeB(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "Verified mold is not present:";
            binSealLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binSealLabel.Size = new System.Drawing.Size(400, 100);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            BinSealGrade.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinSealGrade.Size = new System.Drawing.Size(30, 30);
            BinSealGrade.Validating += BinSealGrade_Validating;

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.Size = new System.Drawing.Size(200, 50);
            firmnessLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Size = new System.Drawing.Size(220, 70);
            this.FirmnessBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label firmLabel = new Label();
            firmLabel.Text = "Firm";
            firmLabel.Location = new System.Drawing.Point(20, 25);
            firmLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(firmLabel);
            this.FirmnessFirm = new RadioButton();
            this.FirmnessFirm.Name = "Firmness Firm";
            this.FirmnessFirm.Location = new System.Drawing.Point(80, 25);
            this.FirmnessFirm.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessFirm);
            this.FirmnessFirm.Checked = true;
            //  Second Button
            Label softLabel = new Label();
            softLabel.Text = "Soft";
            softLabel.Location = new System.Drawing.Point(110, 25);
            softLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(softLabel);
            this.FirmnessSoft = new RadioButton();
            this.FirmnessSoft.Name = "Firmness Soft";
            this.FirmnessSoft.Location = new System.Drawing.Point(170, 25);
            this.FirmnessSoft.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessSoft);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.Size = new System.Drawing.Size(300, 50);
            delvicidLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Size = new System.Drawing.Size(220, 70);
            this.DelvicidBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label trueDelvicidLabel = new Label();
            trueDelvicidLabel.Text = "Yes";
            trueDelvicidLabel.Location = new System.Drawing.Point(20, 25);
            trueDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(trueDelvicidLabel);
            this.DelvicidTrue = new RadioButton();
            this.DelvicidTrue.Name = "Delvicid True";
            this.DelvicidTrue.Location = new System.Drawing.Point(80, 25);
            this.DelvicidTrue.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidTrue);
            this.DelvicidTrue.Checked = true;
            //  Second Button
            Label falseDelvicidLabel = new Label();
            falseDelvicidLabel.Text = "No";
            falseDelvicidLabel.Location = new System.Drawing.Point(120, 25);
            falseDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(falseDelvicidLabel);
            this.DelvicidFalse = new RadioButton();
            this.DelvicidFalse.Name = "Delvicid False";
            this.DelvicidFalse.Location = new System.Drawing.Point(170, 25);
            this.DelvicidFalse.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidFalse);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Size = new System.Drawing.Size(120, 50);
            initialsLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            Initials.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Initials.Size = new System.Drawing.Size(70, 50);
            Initials.CharacterCasing = CharacterCasing.Upper;
            Initials.Validating += Initials_Validating;

            // Add Controls to TableLayoutPanel
            tableLayout.Controls.Add(dateLabel, 0, 0);
            tableLayout.Controls.Add(Date, 1, 0);
            tableLayout.Controls.Add(binWeightLabel, 0, 1);
            tableLayout.Controls.Add(BinWeight, 1, 1);
            tableLayout.Controls.Add(startTimeLabel, 0, 2);
            tableLayout.Controls.Add(StartTime, 1, 2);
            tableLayout.Controls.Add(tempLabel, 0, 3);
            tableLayout.Controls.Add(Temp, 1, 3);
            tableLayout.Controls.Add(firmnessLabel, 0, 4);
            tableLayout.Controls.Add(FirmnessBox, 1, 4);
            tableLayout.Controls.Add(delvicidLabel, 0, 5);
            tableLayout.Controls.Add(DelvicidBox, 1, 5);
            tableLayout.Controls.Add(initialsLabel, 0, 6);
            tableLayout.Controls.Add(Initials, 1, 6);
            tableLayout.Controls.Add(SubmitButton, 1, 8);
            tableLayout.Controls.Add(binSealLabel, 0, 7);
            tableLayout.Controls.Add(BinSealGrade, 1, 7);

            this.Controls.Add(tableLayout);
        }

        private void InitializeBlockTypeC(string productNumber)
        {

            tableLayout = new TableLayoutPanel();
            tableLayout.ColumnCount = 2;
            tableLayout.RowCount = 9;
            tableLayout.Dock = DockStyle.None;  // Remove automatic docking
            tableLayout.AutoSize = true;
            tableLayout.Location = new System.Drawing.Point(300, 250); // Move it right (X=300) and down (Y=200)
            tableLayout.Width = this.Width / 2; // Take up about half the width
            tableLayout.Padding = new Padding(20);
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // Labels
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Controls
            tableLayout.RowStyles.Clear(); // Clear any default row styles

            for (int i = 0; i < tableLayout.RowCount; i++)
            {
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
            }

            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            dateLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(220, 50);
            Date.Text = DateTime.Today.ToString("MM/dd/yyyy");
            this.Date.Validating += new CancelEventHandler(LotDate_Validating_Handler);

            this.piecesNumberLabel = new Label();
            piecesNumberLabel.Text = "Number of Pieces:";
            piecesNumberLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            piecesNumberLabel.Size = new System.Drawing.Size(350, 50);
            this.Controls.Add(piecesNumberLabel);
            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Name = "Number of Pieces";
            NumberPieces.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.NumberPieces.Minimum = 1;
            this.NumberPieces.Maximum = 200;
            this.NumberPieces.Size = new System.Drawing.Size(90, 50);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Weight (lbs.):";
            binWeightLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binWeightLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Name = "Bin Weight";
            BinWeight.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(150, 50);
            BinWeight.Validating += BinWeight_Validating;

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            startTimeLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            StartTime.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(120, 50);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F°):";
            tempLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            tempLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Name = "Temperature";
            Temp.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 80.00M;
            this.Temp.Value = 0.00M;
            this.Temp.Size = new System.Drawing.Size(90, 50);
            Temp.Text = "";
            Temp.Validating += Temp_Validating;

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Size = new System.Drawing.Size(180, 80);
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = ContentAlignment.MiddleCenter;
            SubmitButton.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeC(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "Verified mold is not present:";
            binSealLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binSealLabel.Size = new System.Drawing.Size(400, 100);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            BinSealGrade.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinSealGrade.Size = new System.Drawing.Size(30, 30);
            BinSealGrade.Validating += BinSealGrade_Validating;

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.Size = new System.Drawing.Size(200, 50);
            firmnessLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Size = new System.Drawing.Size(220, 70);
            this.FirmnessBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label firmLabel = new Label();
            firmLabel.Text = "Firm";
            firmLabel.Location = new System.Drawing.Point(20, 25);
            firmLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(firmLabel);
            this.FirmnessFirm = new RadioButton();
            this.FirmnessFirm.Name = "Firmness Firm";
            this.FirmnessFirm.Location = new System.Drawing.Point(80, 25);
            this.FirmnessFirm.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessFirm);
            this.FirmnessFirm.Checked = true;
            //  Second Button
            Label softLabel = new Label();
            softLabel.Text = "Soft";
            softLabel.Location = new System.Drawing.Point(110, 25);
            softLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(softLabel);
            this.FirmnessSoft = new RadioButton();
            this.FirmnessSoft.Name = "Firmness Soft";
            this.FirmnessSoft.Location = new System.Drawing.Point(170, 25);
            this.FirmnessSoft.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessSoft);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.Size = new System.Drawing.Size(300, 50);
            delvicidLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Size = new System.Drawing.Size(220, 70);
            this.DelvicidBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label trueDelvicidLabel = new Label();
            trueDelvicidLabel.Text = "Yes";
            trueDelvicidLabel.Location = new System.Drawing.Point(20, 25);
            trueDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(trueDelvicidLabel);
            this.DelvicidTrue = new RadioButton();
            this.DelvicidTrue.Name = "Delvicid True";
            this.DelvicidTrue.Location = new System.Drawing.Point(80, 25);
            this.DelvicidTrue.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidTrue);
            this.DelvicidTrue.Checked = true;
            //  Second Button
            Label falseDelvicidLabel = new Label();
            falseDelvicidLabel.Text = "No";
            falseDelvicidLabel.Location = new System.Drawing.Point(120, 25);
            falseDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(falseDelvicidLabel);
            this.DelvicidFalse = new RadioButton();
            this.DelvicidFalse.Name = "Delvicid False";
            this.DelvicidFalse.Location = new System.Drawing.Point(170, 25);
            this.DelvicidFalse.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidFalse);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Size = new System.Drawing.Size(120, 50);
            initialsLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            Initials.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Initials.Size = new System.Drawing.Size(70, 50);
            Initials.CharacterCasing = CharacterCasing.Upper;
            Initials.Validating += Initials_Validating;

            // Add Controls to TableLayoutPanel
            tableLayout.Controls.Add(dateLabel, 0, 0);
            tableLayout.Controls.Add(Date, 1, 0);
            tableLayout.Controls.Add(piecesNumberLabel, 0, 1);
            tableLayout.Controls.Add(NumberPieces, 1, 1);
            tableLayout.Controls.Add(binWeightLabel, 0, 2);
            tableLayout.Controls.Add(BinWeight, 1, 2);
            tableLayout.Controls.Add(startTimeLabel, 0, 3);
            tableLayout.Controls.Add(StartTime, 1, 3);
            tableLayout.Controls.Add(tempLabel, 0, 4);
            tableLayout.Controls.Add(Temp, 1, 4);
            tableLayout.Controls.Add(firmnessLabel, 0, 5);
            tableLayout.Controls.Add(FirmnessBox, 1, 5);
            tableLayout.Controls.Add(delvicidLabel, 0, 6);
            tableLayout.Controls.Add(DelvicidBox, 1, 6);
            tableLayout.Controls.Add(initialsLabel, 0, 7);
            tableLayout.Controls.Add(Initials, 1, 7);
            tableLayout.Controls.Add(SubmitButton, 1, 9);
            tableLayout.Controls.Add(binSealLabel, 0, 8);
            tableLayout.Controls.Add(BinSealGrade, 1, 8);

            this.Controls.Add(tableLayout);
        }

        private void InitializeScrap()
        {
            // Define a fixed space between rows (e.g., 10% for spacing)
            float spacePercentage = 10F;

            tableLayout = new TableLayoutPanel();
            tableLayout.ColumnCount = 2;
            tableLayout.RowCount = 9;
            tableLayout.Dock = DockStyle.None;  // Remove automatic docking
            tableLayout.AutoSize = true;
            tableLayout.Location = new System.Drawing.Point(300, 250); // Move it right (X=300) and down (Y=200)
            tableLayout.Width = this.Width / 2; // Take up about half the width
            tableLayout.Padding = new Padding(20);
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // Labels
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Controls
            tableLayout.RowStyles.Clear(); // Clear any default row styles

            for (int i = 0; i < tableLayout.RowCount; i++)
            {
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
            }


            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            dateLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(220, 50);
            Date.Text = DateTime.Today.ToString("MM/dd/yyyy");
            this.Date.Validating += new CancelEventHandler(LotDate_Validating_Handler);

            skidNumberLabel = new Label();
            skidNumberLabel.Text = "Skid/Tote Number:";
            skidNumberLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            skidNumberLabel.Size = new System.Drawing.Size(350, 50);
            this.Controls.Add(skidNumberLabel);
            ToteSkidNumber = new NumericUpDown();
            ToteSkidNumber.Name = "Skid Number";
            ToteSkidNumber.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            ToteSkidNumber.Size = new System.Drawing.Size(90, 50);
            ToteSkidNumber.Maximum = 199;
            ToteSkidNumber.Minimum = 0;
            ToteSkidNumber.Value = 0;
            ToteSkidNumber.Text = "";
            ToteSkidNumber.Validating += ToteSkidNumber_Validating;

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Bin Weight (lbs.):";
            binWeightLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binWeightLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Name = "Bin Weight";
            BinWeight.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(150, 50);
            BinWeight.Validating += BinWeight_Validating;

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            startTimeLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            StartTime.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(120, 50);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F°):";
            tempLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            tempLabel.Size = new System.Drawing.Size(250, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Name = "Temperature";
            Temp.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 80.00M;
            this.Temp.Value = 0.00M;
            this.Temp.Size = new System.Drawing.Size(90, 50);
            Temp.Text = "";
            Temp.Validating += Temp_Validating;

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Size = new System.Drawing.Size(180, 80);
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = ContentAlignment.MiddleCenter;
            SubmitButton.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedScrap(sender, e); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "Verified mold is not present:";
            binSealLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            binSealLabel.Size = new System.Drawing.Size(400, 100);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            BinSealGrade.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BinSealGrade.Size = new System.Drawing.Size(30, 30);
            BinSealGrade.Validating += BinSealGrade_Validating;

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.Size = new System.Drawing.Size(200, 50);
            firmnessLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Size = new System.Drawing.Size(220, 70);
            this.FirmnessBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label firmLabel = new Label();
            firmLabel.Text = "Firm";
            firmLabel.Location = new System.Drawing.Point(20, 25);
            firmLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(firmLabel);
            this.FirmnessFirm = new RadioButton();
            this.FirmnessFirm.Name = "Firmness Firm";
            this.FirmnessFirm.Location = new System.Drawing.Point(80, 25);
            this.FirmnessFirm.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessFirm);
            this.FirmnessFirm.Checked = true;
            //  Second Button
            Label softLabel = new Label();
            softLabel.Text = "Soft";
            softLabel.Location = new System.Drawing.Point(110, 25);
            softLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(softLabel);
            this.FirmnessSoft = new RadioButton();
            this.FirmnessSoft.Name = "Firmness Soft";
            this.FirmnessSoft.Location = new System.Drawing.Point(170, 25);
            this.FirmnessSoft.Size = new System.Drawing.Size(30, 30);
            this.FirmnessBox.Controls.Add(this.FirmnessSoft);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.Size = new System.Drawing.Size(300, 50);
            delvicidLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Size = new System.Drawing.Size(220, 70);
            this.DelvicidBox.FlatStyle = FlatStyle.Standard;
            //  First Button
            Label trueDelvicidLabel = new Label();
            trueDelvicidLabel.Text = "Yes";
            trueDelvicidLabel.Location = new System.Drawing.Point(20, 25);
            trueDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(trueDelvicidLabel);
            this.DelvicidTrue = new RadioButton();
            this.DelvicidTrue.Name = "Delvicid True";
            this.DelvicidTrue.Location = new System.Drawing.Point(80, 25);
            this.DelvicidTrue.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidTrue);
            this.DelvicidTrue.Checked = true;
            //  Second Button
            Label falseDelvicidLabel = new Label();
            falseDelvicidLabel.Text = "No";
            falseDelvicidLabel.Location = new System.Drawing.Point(120, 25);
            falseDelvicidLabel.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(falseDelvicidLabel);
            this.DelvicidFalse = new RadioButton();
            this.DelvicidFalse.Name = "Delvicid False";
            this.DelvicidFalse.Location = new System.Drawing.Point(170, 25);
            this.DelvicidFalse.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidFalse);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Size = new System.Drawing.Size(120, 50);
            initialsLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            Initials.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Initials.Size = new System.Drawing.Size(70, 50);
            Initials.CharacterCasing = CharacterCasing.Upper;
            Initials.Validating += Initials_Validating;

            // Add Controls to TableLayoutPanel
            tableLayout.Controls.Add(dateLabel, 0, 0);
            tableLayout.Controls.Add(Date, 1, 0);
            tableLayout.Controls.Add(skidNumberLabel, 0, 1);
            tableLayout.Controls.Add(ToteSkidNumber, 1, 1);
            tableLayout.Controls.Add(binWeightLabel, 0, 2);
            tableLayout.Controls.Add(BinWeight, 1, 2);
            tableLayout.Controls.Add(startTimeLabel, 0, 3);
            tableLayout.Controls.Add(StartTime, 1, 3);
            tableLayout.Controls.Add(tempLabel, 0, 4);
            tableLayout.Controls.Add(Temp, 1, 4);
            tableLayout.Controls.Add(firmnessLabel, 0, 5);
            tableLayout.Controls.Add(FirmnessBox, 1, 5);
            tableLayout.Controls.Add(delvicidLabel, 0, 6);
            tableLayout.Controls.Add(DelvicidBox, 1, 6);
            tableLayout.Controls.Add(initialsLabel, 0, 7);
            tableLayout.Controls.Add(Initials, 1, 7);
            tableLayout.Controls.Add(SubmitButton, 1, 9);
            tableLayout.Controls.Add(binSealLabel, 0, 8);
            tableLayout.Controls.Add(BinSealGrade, 1, 8);

            // Add TableLayoutPanel to Form
            this.Controls.Add(tableLayout);
        }

        private void InitializePowder()
        {
            // Define a fixed space between rows (e.g., 10% for spacing)
            float spacePercentage = 10F;

            tableLayout = new TableLayoutPanel();
            tableLayout.ColumnCount = 2;
            tableLayout.RowCount = 4;
            tableLayout.Dock = DockStyle.None;  // Remove automatic docking
            tableLayout.AutoSize = true;
            tableLayout.Location = new System.Drawing.Point(300, 250); // Move it right (X=300) and down (Y=200)
            tableLayout.Width = this.Width / 2; // Take up about half the width
            tableLayout.Padding = new Padding(20);
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // Labels
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F)); // Controls
            tableLayout.RowStyles.Clear(); // Clear any default row styles

            for (int i = 0; i < tableLayout.RowCount; i++)
            {
                tableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F));
            }

            this.bagCountLabel = new Label();
            bagCountLabel.Text = "Bag Count:";
            bagCountLabel.Size = new System.Drawing.Size(200, 50);
            bagCountLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.BagCount = new NumericUpDown();
            BagCount.Name = "Bag Count";
            BagCount.Size = new System.Drawing.Size(70, 50);
            BagCount.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            BagCount.Maximum = 10;
            BagCount.Minimum = 0;
            BagCount.Increment = 1;
            BagCount.Value = 0;
            BagCount.Text = "";
            BagCount.Validating += BagCount_Validating;

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            startTimeLabel.Size = new System.Drawing.Size(200, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            StartTime.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(120, 50);

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Size = new System.Drawing.Size(120, 50);
            initialsLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            Initials.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.Initials.Size = new System.Drawing.Size(70, 50);
            Initials.CharacterCasing = CharacterCasing.Upper;
            Initials.Validating += Initials_Validating;

            this.powderLotNumberLabel = new Label();
            powderLotNumberLabel.Text = "Lot Number:";
            powderLotNumberLabel.Size = new System.Drawing.Size(200, 60);
            powderLotNumberLabel.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            this.PowderLotNumber = new TextBox();
            PowderLotNumber.Name = "Powder Lot Number";
            PowderLotNumber.Size = new System.Drawing.Size(200, 30);
            PowderLotNumber.Font = new System.Drawing.Font("Arial", 14, FontStyle.Regular);
            PowderLotNumber.CharacterCasing = CharacterCasing.Upper;
            PowderLotNumber.Validating += PowderLotNumber_Validating;

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Size = new System.Drawing.Size(180, 80);
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = ContentAlignment.MiddleCenter;
            SubmitButton.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold);
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedPowder(sender, e); };

            // Add Controls to TableLayoutPanel
            tableLayout.Controls.Add(bagCountLabel, 0, 0);
            tableLayout.Controls.Add(BagCount, 1, 0);
            tableLayout.Controls.Add(startTimeLabel, 0, 1);
            tableLayout.Controls.Add(StartTime, 1, 1);
            tableLayout.Controls.Add(initialsLabel, 0, 2);
            tableLayout.Controls.Add(Initials, 1, 2);
            tableLayout.Controls.Add(powderLotNumberLabel, 0, 3);
            tableLayout.Controls.Add(PowderLotNumber, 1, 3);
            tableLayout.Controls.Add(SubmitButton, 1, 4);


            // Add TableLayoutPanel to Form
            this.Controls.Add(tableLayout);
        }


        //      SUBMITTING AND WRITING METHODS

        public static string ColumnNumberToName(int columnNumber)
        {
            string columnName = String.Empty;
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        public bool InitializeSubmitCheck(object data)
        {
            bool isAllValid = this.ValidateChildren();
            if (!isAllValid) { return true; }

            var result = CustomMessageBox.Show(data);

            if (result == DialogResult.Yes)
            {
                MessageBox.Show("Item was successfully tracked");
                return false;
            }
            else
            {
                return true;
            }
        }

        private void SubmitButton_ClickedTypeA(object sender, EventArgs e, string productNumber)
        {

            var data = new[]
                {
                    new { Column1 = productNumber, Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = "#" + ToteSkidNumber.Value.ToString(), Column2 = NumberPieces.Value.ToString() + " pcs" },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "°F" , Column2 = "GOOD SEAL"},
                    new { Column1 = Initials.Text , Column2 = "" }
                };

            if (InitializeSubmitCheck(data))
            {
                return;
            }

            this.ws = wb.Worksheet(productNumber);

            bool flag = false;
            int columnNumber = 3;

            //Check for Date
            //
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value = this.Date.Value.Date;
                    flag = true;
                }
                else if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value.GetDateTime().Date == this.Date.Value.Date)
                {
                    flag = true;
                }
                else
                {
                    columnNumber += 10;
                }
            }

            columnNumber -= 1;
            int rowNumber = 4;
            flag = false;
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.ToteSkidNumber.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.NumberPieces.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.BinWeight.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.StartTime.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.Temp.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "GOOD";
                    columnNumber += 1;
                    if (this.FirmnessFirm.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "F";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "S";
                    }
                    columnNumber += 1;
                    if (this.DelvicidTrue.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "YES";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "NO";
                    }
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = Initials.Text;
                    flag = true;
                }
                else
                {
                    rowNumber += 1;
                }
            }

            string i = "";
            if (ComboBox1.Text.ToString() == "008-000005 PS Purchased" || ComboBox1.Text.ToString() == "008-000021 WM Purchased" || ComboBox1.Text.ToString() == "002-000035 Scrap")
            {
                string itemText = TrimItemNumber(ComboBox1.Text.ToString());
                i = itemText + "    Tote/Bin #: " + ToteSkidNumber.Value.ToString() + "    Lot: " + Date.Value.ToShortDateString() + "    Time: " + StartTime.Value.ToShortTimeString()
                        + "    ID: " + Initials.Text.ToString();

            }
            else
            {
                i = ComboBox1.Text.ToString() + "    Tote/Bin #: " + ToteSkidNumber.Value.ToString() + "    Lot: " + Date.Value.ToShortDateString() + "    Time: " + StartTime.Value.ToShortTimeString()
                        + "    ID: " + Initials.Text.ToString();
            }
            UpdateList(i);

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedTypeB(object sender, EventArgs e, string productNumber)
        {
            var data = new[]
                {
                    new { Column1 = productNumber, Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "°F" , Column2 = "GOOD SEAL"},
                    new { Column1 = Initials.Text , Column2 = "" }
                };

            if (InitializeSubmitCheck(data))
            {
                return;
            }

            this.ws = wb.Worksheet(productNumber);

            bool flag = false;
            int columnNumber = 3;

            //Check for Date
            //
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value = this.Date.Value.Date;
                    flag = true;
                }
                else if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value.GetDateTime().Date == this.Date.Value.Date)
                {
                    flag = true;
                }
                else
                {
                    columnNumber += 8;
                }
            }

            columnNumber -= 1;
            int rowNumber = 4;
            flag = false;
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.BinWeight.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.StartTime.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.Temp.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "GOOD";
                    columnNumber += 1;
                    if (this.FirmnessFirm.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "F";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "S";
                    }
                    columnNumber += 1;
                    if (this.DelvicidTrue.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "YES";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "NO";
                    }
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = Initials.Text;
                    flag = true;
                }
                else
                {
                    rowNumber += 1;
                }
            }
            string itemText = TrimItemNumber(ComboBox1.Text.ToString());
            string i = itemText + "    Lot: " + Date.Value.ToShortDateString() + "    Time: " + StartTime.Value.ToShortTimeString()
                    + "    ID: " + Initials.Text.ToString();
            UpdateList(i);

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedTypeC(object sender, EventArgs e, string productNumber)
        {
            var data = new[]
                {
                    new { Column1 = productNumber, Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "°F" , Column2 = "GOOD SEAL"},
                    new { Column1 = NumberPieces.Value.ToString() + " pcs" , Column2 = Initials.Text }
                };

            if (InitializeSubmitCheck(data))
            {
                return;
            }

            this.ws = wb.Worksheet(productNumber);

            bool flag = false;
            int columnNumber = 3;

            //Check for Date
            //
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value = this.Date.Value.Date;
                    flag = true;
                }
                else if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value.GetDateTime().Date == this.Date.Value.Date)
                {
                    flag = true;
                }
                else
                {
                    columnNumber += 10;
                }
            }

            columnNumber -= 1;
            int rowNumber = 4;
            flag = false;
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "0";
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.NumberPieces.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.BinWeight.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.StartTime.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.Temp.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "GOOD";
                    columnNumber += 1;
                    if (this.FirmnessFirm.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "F";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "S";
                    }
                    columnNumber += 1;
                    if (this.DelvicidTrue.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "YES";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "NO";
                    }
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = Initials.Text;
                    flag = true;
                }
                else
                {
                    rowNumber += 1;
                }
            }

            string itemText = TrimItemNumber(ComboBox1.Text.ToString());
            string i = itemText + "    Lot: " + Date.Value.ToShortDateString() + "    Time: " + StartTime.Value.ToShortTimeString()
                    + "    ID: " + Initials.Text.ToString();
            UpdateList(i);

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedScrap(object sender, EventArgs e)
        {
            var data = new[]
                {
                    new { Column1 = "Scrap Tote", Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "°F" , Column2 = "GOOD SEAL"},
                    new { Column1 = "#" + ToteSkidNumber.Value.ToString() , Column2 = Initials.Text }
                };

            if (InitializeSubmitCheck(data))
            {
                return;
            }

            this.ws = wb.Worksheet("All Scrap");

            bool flag = false;
            int columnNumber = 3;

            //Check for Date
            //
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value = this.Date.Value.Date;
                    flag = true;
                }
                else if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + "2").Value.GetDateTime().Date == this.Date.Value.Date)
                {
                    flag = true;
                }
                else
                {
                    columnNumber += 10;
                }
            }

            columnNumber -= 1;
            int rowNumber = 4;
            flag = false;
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.ToteSkidNumber.Value;
                    columnNumber += 2;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.BinWeight.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.StartTime.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.Temp.Value;
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "GOOD";
                    columnNumber += 1;
                    if (this.FirmnessFirm.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "F";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "S";
                    }
                    columnNumber += 1;
                    if (this.DelvicidTrue.Checked)
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "YES";
                    }
                    else
                    {
                        ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = "NO";
                    }
                    columnNumber += 1;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = Initials.Text;
                    flag = true;
                }
                else
                {
                    rowNumber += 1;
                }
            }

            string i = "Scrap" + "    Lot: " + Date.Value.ToShortDateString() + "    Weight: " + BinWeight.Value.ToString() + "lbs.    Time: " + StartTime.Value.ToShortTimeString()
                    + "    ID: " + Initials.Text.ToString();
            UpdateList(i);

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedPowder(object sender, EventArgs e)
        {
            var data = new[]
                {
                    new { Column1 = "Powder", Column2 = BagCount.Value.ToString() + " Bag(s)" },
                    new { Column1 = PowderLotNumber.Text  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Initials.Text.ToString(), Column2 = ""},
                };

            if (InitializeSubmitCheck(data))
            {
                return;
            }

            this.ws = wb.Worksheet("Powder");

            if (ws.Worksheet.Cell("C2").IsEmpty())
            {
                ws.Worksheet.Cell("C2").Value = DateTime.Now.Date;
            }

            int columnNumber = 2;
            int rowNumber = 4;

            bool flag = false;
            while (flag == false)
            {
                if (ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).IsEmpty())
                {
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.BagCount.Value;
                    columnNumber += 3;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.StartTime.Value;
                    columnNumber++;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.PowderLotNumber.Text;
                    columnNumber++;
                    ws.Worksheet.Cell(ColumnNumberToName(columnNumber) + rowNumber.ToString()).Value = this.Initials.Text;
                    flag = true;
                }
                else
                {
                    rowNumber++;
                }
            }

            string i = ComboBox1.Text.ToString() + "    Bag Count: " + BagCount.Value.ToString() + "    Lot #: " + PowderLotNumber.Text.ToString() + "    Time: " + StartTime.Value.ToShortTimeString()
                    + "    ID: " + Initials.Text.ToString();
            UpdateList(i);

            NewSelection();
            this.wb.Save();
        }


        //      VALIDATION METHODS

        private void LotDate_Validating(object? sender, CancelEventArgs e, string productNumber = "")
        {

            DateTime today = DateTime.Today;
            DateTime maxValidDate = today.AddDays(-45);
            DateTime minApprovalDate = string.Equals(productNumber, "scrap", StringComparison.OrdinalIgnoreCase)
                                        ? today.AddDays(-25)
                                        : today.AddDays(-30);

            if (this.Date.Value < maxValidDate)
            {
                MessageBox.Show("This item's lot date is over 45 days old. Please enter a valid lot date or contact your manager.");
                e.Cancel = true;
                errorProvider.SetError(Date, "Please Enter a Valid Date");
            }
            else if (this.Date.Value < minApprovalDate)
            {
                MessageBox.Show($"This item's lot date requires manager approval. Please confirm before proceeding.");
                errorProvider.SetError(Date, "");
                e.Cancel = false; // Allows continuation but with a warning
            }
            else if (this.Date.Value == today)
            {
                MessageBox.Show("This item's lot date is today. Please enter a valid lot date or contact your manager.");
                e.Cancel = true;
                errorProvider.SetError(Date, "Please Enter a Valid Date");
            }
            else if (this.Date.Value > today)
            {
                MessageBox.Show("This item's lot date is in the future. Please enter a valid lot date or contact your manager.");
                e.Cancel = true;
                errorProvider.SetError(Date, "Please Enter a Valid Date");
            }
            else
            {
                ClearValidationError(Date, e);
            }


        }
        // Wrapper method to match CancelEventHandler signature
        private void LotDate_Validating_Handler(object sender, CancelEventArgs e)
        {
            if (ComboBox1.Text == "002-000035 Scrap")
            {
                LotDate_Validating(sender, e, "scrap");
            }
            else
            {
                LotDate_Validating(sender, e, ""); // Calls the main method with an empty product number
            }
        }

        private void ToteSkidNumber_Validating(object? sender, CancelEventArgs e)
        {

            if (this.ToteSkidNumber.Value == 0 || this.ToteSkidNumber.Text == "")
            {
                MessageBox.Show("Please enter a valid skid/tote number. Number cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(ToteSkidNumber, "Please Enter a Valid Tote/Skid Number");
            }
            else
            {
                ClearValidationError(ToteSkidNumber, e);
            }
        }

        private void BagCount_Validating(object? sender, CancelEventArgs e)
        {

            if (this.BagCount.Value == 0 || this.BagCount.Text == "")
            {
                MessageBox.Show("Please enter a valid bag count. Number cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(BagCount, "Please Enter a Valid Bag Count");
            }
            else
            {
                ClearValidationError(BagCount, e);
            }
        }

        private void PowderLotNumber_Validating(object? sender, CancelEventArgs e)
        {

            if (this.PowderLotNumber.Text.Length < 1 || this.PowderLotNumber.Text.Length > 14)
            {
                MessageBox.Show("Please enter a valid powder lot number.");
                e.Cancel = true;
                errorProvider.SetError(PowderLotNumber, "Please Enter a Valid Lot");
            }
            else if (Regex.IsMatch(PowderLotNumber.Text, @"[^a-zA-Z0-9]"))
            {
                MessageBox.Show("Lot number cannot contain special characters.");
                e.Cancel = true;
                errorProvider.SetError(PowderLotNumber, "Lot number cannot contain special characters.");
            }
            else
            {
                ClearValidationError(PowderLotNumber, e);
            }
        }

        private void BinWeight_Validating(object? sender, CancelEventArgs e)
        {

            if (this.BinWeight.Value <= 0.00M || this.BinWeight.Text == "")
            {
                MessageBox.Show("Please enter a valid weight. Weight cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(BinWeight, "Please Enter a Valid Weight");
            }
            else if (BinWeight.Value >= 1200M)
            {
                MessageBox.Show("Please enter a valid weight. Weight cannot be over 1200lbs.");
                e.Cancel = true;
                errorProvider.SetError(BinWeight, "Please Enter a Valid Weight. Weight cannot be over 1200lbs.");
            }
            else
            {
                ClearValidationError(BinWeight, e);
            }
        }

        private void Temp_Validating(object? sender, CancelEventArgs e)
        {

            if (Temp.Value == 0.00M || Temp.Text == "")
            {
                MessageBox.Show("Please enter a valid temperature.");
                e.Cancel = true;
                errorProvider.SetError(Temp, "Please Enter a Valid Temperature");
            }
            else if (Temp.Value >= 60.00M)
            {
                MessageBox.Show("Temperature is above 60°F. Please notify a manager to destory or put away item.");
                e.Cancel = true;
                errorProvider.SetError(Temp, "Temperature is Too High");
            }
            else if (Temp.Value >= 47.00M)
            {
                MessageBox.Show("This item's temperature requires manager approval. Temperature is above 47°F.");
                e.Cancel = false;
                ClearValidationError(Temp, e);
            }
            else
            {
                ClearValidationError(Temp, e);
            }
        }

        private void BinSealGrade_Validating(object? sender, CancelEventArgs e)
        {

            if (BinSealGrade.Checked == false)
            {
                MessageBox.Show("Please mark that you have verified no mold is present.");
                e.Cancel = true;
                errorProvider.SetError(BinSealGrade, "Please Check to Verify No Mold is Present");
            }
            else
            {
                ClearValidationError(BinSealGrade, e);
            }
        }

        private void Initials_Validating(object? sender, CancelEventArgs e)
        {

            if (Initials.Text.Length < 2 || Initials.Text.Length > 3)
            {
                MessageBox.Show("Initials must be 2-3 letters in length.");
                e.Cancel = true;
                errorProvider.SetError(Initials, "Initials must be 2-3 letters in length");
            }
            else if (Regex.IsMatch(Initials.Text, @"\d") || Regex.IsMatch(Initials.Text, @"[^a-zA-Z0-9]"))
            {
                MessageBox.Show("Initials cannot contain numbers or special characters.");
                e.Cancel = true;
                errorProvider.SetError(Initials, "Initials cannot contain numbers or special characters.");
            }
            else
            {
                ClearValidationError(Initials, e);
            }
        }

        private void ClearValidationError(System.Windows.Forms.Control control, CancelEventArgs e)
        {
            e.Cancel = false;
            errorProvider.SetError(control, "");
        }

        private void UpdateList(string item)
        {
            if (!string.IsNullOrEmpty(item))
            {
                if (L.Count == 10)
                {
                    L.RemoveAt(0); // Remove oldest entry
                }

                L.Add(item);

                rightPanel.Controls.Clear();

                foreach (string i in L)
                {
                    Label label = new Label();
                    label.Text = i;
                    label.ForeColor = System.Drawing.Color.Black;
                    label.BackColor = System.Drawing.Color.White;
                    label.Font = new System.Drawing.Font("Arial", 14, FontStyle.Bold); // Bigger text
                    label.AutoSize = true;
                    rightPanel.Controls.Add(label);
                }
            }
        }

        private string TrimItemNumber(string item)
        {
            return item.Substring(11);
        }
    }
}
