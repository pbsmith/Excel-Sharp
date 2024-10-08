using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Globalization;
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
            this.solutionDirectory = GetSolutionDirectoryInfo().ToString().Remove(GetSolutionDirectoryInfo().ToString().Length - 18);
            Trace.WriteLine(solutionDirectory);
            string yearDirectory = solutionDirectory + thisYear.ToString();
            if(!Directory.Exists(yearDirectory))
            {
                Directory.CreateDirectory(yearDirectory);
            }
            string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(thisMonth);
            string monthDirectory = Path.Combine(yearDirectory, monthName);
            if (!Directory.Exists(monthDirectory))
            {
                Directory.CreateDirectory(monthDirectory);
            }
            string filePath = Path.Combine(monthDirectory, thisDay.ToString() + "-" + monthName + "_Shred_Usage_Output" + ".xlsx");
            ogWorkbook = Path.Combine(solutionDirectory, "blank.xlsx");
            if (!File.Exists(filePath))
            {
                Trace.WriteLine(ogWorkbook);
                File.Copy(ogWorkbook, filePath);
                
            }
            this.wb = new XLWorkbook(filePath);
            InitializeComponent();
            this.Text = "Miceli Dairy Products - Block Usage Reporting Tool";
            this.ShowIcon = false;
            InitializeComboBox();
            //this.wb = new XLWorkbook(@"C:\Users\psmith\workspace\Excel-Sharp\blank.xlsx");
            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            this.Width = screenWidth / 2;
            this.Height = Screen.PrimaryScreen.WorkingArea.Height;
        }

        //      ERROR PROVIDER
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
            this.ComboBox1.Location = new System.Drawing.Point(100, 38);
            this.ComboBox1.Name = "ComboBox1";
            this.ComboBox1.Size = new System.Drawing.Size(200, 50);
            this.ComboBox1.TabIndex = 0;
            this.ComboBox1.Text = "Select Block Item";
            string[] installs = new string[] { "001-000133", "001-000169", "001-000195", "001-000229", "001-000360", "001-000455", "001-000470", "001-000508",
            "001-000525", "001-000528", "008-000005 PS Purchased", "008-000021 WM Purchased", "008-000001 Asiago", "008-000002 Cheddar", "008-000006 Parmesan",
            "008-000010 White Ched", "008-000022 Meunster", "008-000007 Provolone", "002-000035 Scrap", "Powder"};
            ComboBox1.Items.AddRange(installs);
            this.Controls.Add(this.ComboBox1);

            // Hook up the event handler.
            this.ComboBox1.SelectedIndexChanged +=
                new System.EventHandler(ComboBox1_SelectedIndexChanged);
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
                case "001-000508":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                     001-000508      ";
                    InitializeBlockTypeA("1-508");
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
            this.Controls.Remove(ToteSkidNumber);
            if (ToteSkidNumber != null) { ToteSkidNumber.Dispose(); }
            this.Controls.Remove(BinWeight);
            if (BinWeight != null) { BinWeight.Dispose(); }
            this.Controls.Remove(Temp);
            if (Temp != null) { Temp.Dispose(); }
            this.Controls.Remove(DelvicidBox);
            if (DelvicidBox != null) { DelvicidBox.Dispose(); }
            this.Controls.Remove(FirmnessBox);
            if (FirmnessBox != null) { FirmnessBox.Dispose(); }
            this.Controls.Remove(firmnessLabel);
            this.Controls.Remove(delvicidLabel);
            this.Controls.Remove(Date);
            if (Date != null) { Date.Dispose(); }
            this.Controls.Remove(NumberPieces);
            if (NumberPieces != null) { NumberPieces.Dispose(); }
            this.Controls.Remove(StartTime);
            if (StartTime != null) { StartTime.Dispose(); }
            this.Controls.Remove(tempLabel);
            this.Controls.Remove(dateLabel);
            this.Controls.Remove(skidNumberLabel);
            this.Controls.Remove(piecesNumberLabel);
            this.Controls.Remove(binWeightLabel);
            this.Controls.Remove(startTimeLabel);
            this.Controls.Remove(binSealLabel);
            this.Controls.Remove(BinSealGrade);
            if (BinSealGrade != null) { BinSealGrade.Dispose(); }
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
        }

        private void InitializeBlockTypeA(string productNumber)
        {
            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.TextAlign = ContentAlignment.MiddleRight;
            dateLabel.Location = new System.Drawing.Point(140, 218);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Location = new System.Drawing.Point(250, 218);
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(140, 50);
            Date.Text = "01/01/2024";
            Date.Validating += LotDate_Validating;
            this.Controls.Add(Date);

            skidNumberLabel = new Label();
            skidNumberLabel.Text = "Skid/Tote Number:";
            skidNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            skidNumberLabel.Location = new System.Drawing.Point(100, 288);
            skidNumberLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(skidNumberLabel);
            ToteSkidNumber = new NumericUpDown();
            ToteSkidNumber.Location = new System.Drawing.Point(250, 298);
            ToteSkidNumber.Name = "Skid Number";
            ToteSkidNumber.Size = new System.Drawing.Size(70, 50);
            ToteSkidNumber.Maximum = 199;
            ToteSkidNumber.Minimum = 0;
            ToteSkidNumber.Value = 0;
            ToteSkidNumber.Text = "";
            ToteSkidNumber.Validating += ToteSkidNumber_Validating;
            this.Controls.Add(ToteSkidNumber);

            this.piecesNumberLabel = new Label();
            piecesNumberLabel.Text = "Number of Pieces:";
            piecesNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            piecesNumberLabel.Location = new System.Drawing.Point(145, 368);
            piecesNumberLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(piecesNumberLabel);
            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Location = new System.Drawing.Point(250, 378);
            this.NumberPieces.Name = "Number of Pieces";
            this.NumberPieces.Minimum = 1;
            this.NumberPieces.Maximum = 200;
            this.NumberPieces.Value = 160;
            this.NumberPieces.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.NumberPieces);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Bin Weight (lbs.):";
            binWeightLabel.TextAlign = ContentAlignment.MiddleRight;
            binWeightLabel.Location = new System.Drawing.Point(100, 448);
            binWeightLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Location = new System.Drawing.Point(250, 458);
            this.BinWeight.Name = "Bin Weight";
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            BinWeight.Validating += BinWeight_Validating;
            this.Controls.Add(this.BinWeight);

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.TextAlign = ContentAlignment.MiddleRight;
            startTimeLabel.Location = new System.Drawing.Point(150, 528);
            startTimeLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F�):";
            tempLabel.TextAlign = ContentAlignment.MiddleRight;
            tempLabel.Location = new System.Drawing.Point(120, 608);
            tempLabel.Size = new System.Drawing.Size(120, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(250, 618);
            this.Temp.Name = "Temperature";
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 50.00M;
            this.Temp.Value = 0.00M;
            this.Temp.Size = new System.Drawing.Size(100, 50);
            Temp.Text = "";
            Temp.Validating += Temp_Validating;
            this.Controls.Add(this.Temp);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 748);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeA(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "       Bin Seal:                           (By checking this box you confirm that the bin is sealed adequately)";
            binSealLabel.TextAlign = ContentAlignment.MiddleRight;
            binSealLabel.Location = new System.Drawing.Point(460, 188);
            binSealLabel.Size = new System.Drawing.Size(200, 140);
            this.Controls.Add(binSealLabel);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(700, 228);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            BinSealGrade.Validating += BinSealGrade_Validating;
            this.Controls.Add(this.BinSealGrade);

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.TextAlign = ContentAlignment.MiddleRight;
            firmnessLabel.Location = new System.Drawing.Point(500, 393);
            firmnessLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(firmnessLabel);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Location = new System.Drawing.Point(620, 378);
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
            this.Controls.Add(this.FirmnessBox);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.TextAlign = ContentAlignment.MiddleRight;
            delvicidLabel.Location = new System.Drawing.Point(500, 463);
            delvicidLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(delvicidLabel);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Location = new System.Drawing.Point(620, 448);
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
            this.Controls.Add(this.DelvicidBox);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Location = new System.Drawing.Point(580, 578);
            initialsLabel.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            this.Initials.Location = new System.Drawing.Point(660, 578);
            this.Initials.Size = new System.Drawing.Size(40, 30);
            Initials.Validating += Initials_Validating;
            this.Controls.Add(this.Initials);
        }

        private void InitializeBlockTypeB(string productNumber)
        {
            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.TextAlign = ContentAlignment.MiddleRight;
            dateLabel.Location = new System.Drawing.Point(140, 218);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Location = new System.Drawing.Point(250, 218);
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(140, 50);
            Date.Text = "01/01/2024";
            Date.Validating += LotDate_Validating;
            this.Controls.Add(Date);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Block Weight (lbs.):";
            binWeightLabel.TextAlign = ContentAlignment.MiddleRight;
            binWeightLabel.Location = new System.Drawing.Point(100, 448);
            binWeightLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Location = new System.Drawing.Point(250, 458);
            this.BinWeight.Name = "Block Weight";
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            BinWeight.Validating += BinWeight_Validating;
            this.Controls.Add(this.BinWeight);

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.TextAlign = ContentAlignment.MiddleRight;
            startTimeLabel.Location = new System.Drawing.Point(150, 528);
            startTimeLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F�):";
            tempLabel.TextAlign = ContentAlignment.MiddleRight;
            tempLabel.Location = new System.Drawing.Point(120, 608);
            tempLabel.Size = new System.Drawing.Size(120, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(250, 618);
            this.Temp.Name = "Temperature";
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 50.00M;
            this.Temp.Value = 0.00M;
            Temp.Text = "";
            this.Temp.Size = new System.Drawing.Size(100, 50);
            Temp.Validating += Temp_Validating;
            this.Controls.Add(this.Temp);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 748);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeB(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "       Bin Seal:                           (By checking this box you confirm that the bin is sealed adequately)";
            binSealLabel.TextAlign = ContentAlignment.MiddleRight;
            binSealLabel.Location = new System.Drawing.Point(460, 188);
            binSealLabel.Size = new System.Drawing.Size(200, 140);
            this.Controls.Add(binSealLabel);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(700, 228);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            BinSealGrade.Validating += BinSealGrade_Validating;
            this.Controls.Add(this.BinSealGrade);

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.TextAlign = ContentAlignment.MiddleRight;
            firmnessLabel.Location = new System.Drawing.Point(500, 393);
            firmnessLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(firmnessLabel);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Location = new System.Drawing.Point(620, 378);
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
            this.Controls.Add(this.FirmnessBox);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.TextAlign = ContentAlignment.MiddleRight;
            delvicidLabel.Location = new System.Drawing.Point(500, 463);
            delvicidLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(delvicidLabel);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Location = new System.Drawing.Point(620, 448);
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
            this.Controls.Add(this.DelvicidBox);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Location = new System.Drawing.Point(580, 578);
            initialsLabel.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            this.Initials.Location = new System.Drawing.Point(660, 578);
            this.Initials.Size = new System.Drawing.Size(40, 30);
            this.Initials.Validating += Initials_Validating;
            this.Controls.Add(this.Initials);
        }

        private void InitializeBlockTypeC(string productNumber)
        {
            dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.TextAlign = ContentAlignment.MiddleRight;
            dateLabel.Location = new System.Drawing.Point(140, 218);
            this.Controls.Add(dateLabel);
            Date = new DateTimePicker();
            Date.Location = new System.Drawing.Point(250, 218);
            Date.Name = "Date Picker";
            Date.CustomFormat = "MM-dd-yyyy";
            Date.Format = DateTimePickerFormat.Custom;
            Date.Size = new System.Drawing.Size(140, 50);
            Date.Text = "01/01/2024";
            Date.Validating += LotDate_Validating;
            this.Controls.Add(Date);

            //Case Count Below

            this.piecesNumberLabel = new Label();
            piecesNumberLabel.Text = "Number of Pieces Used:";
            piecesNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            piecesNumberLabel.Location = new System.Drawing.Point(145, 368);
            piecesNumberLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(piecesNumberLabel);
            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Location = new System.Drawing.Point(250, 378);
            this.NumberPieces.Name = "Number of Pieces";
            this.NumberPieces.Minimum = 1;
            this.NumberPieces.Maximum = 100;
            this.NumberPieces.Text = "";
            this.NumberPieces.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.NumberPieces);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Weight (lbs.):";
            binWeightLabel.TextAlign = ContentAlignment.MiddleRight;
            binWeightLabel.Location = new System.Drawing.Point(100, 448);
            binWeightLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Location = new System.Drawing.Point(250, 458);
            this.BinWeight.Name = "Bin Weight";
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            BinWeight.Validating += BinWeight_Validating;
            this.Controls.Add(this.BinWeight);

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.TextAlign = ContentAlignment.MiddleRight;
            startTimeLabel.Location = new System.Drawing.Point(150, 528);
            startTimeLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F�):";
            tempLabel.TextAlign = ContentAlignment.MiddleRight;
            tempLabel.Location = new System.Drawing.Point(120, 608);
            tempLabel.Size = new System.Drawing.Size(120, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(250, 618);
            this.Temp.Name = "Temperature";
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 50.00M;
            this.Temp.Value = 0.00M;
            Temp.Text = "";
            this.Temp.Size = new System.Drawing.Size(100, 50);
            this.Temp.Validating += Temp_Validating;
            this.Controls.Add(this.Temp);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 748);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedTypeC(sender, e, productNumber); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "       Bin Seal:                           (By checking this box you confirm that the bin is sealed adequately)";
            binSealLabel.TextAlign = ContentAlignment.MiddleRight;
            binSealLabel.Location = new System.Drawing.Point(460, 188);
            binSealLabel.Size = new System.Drawing.Size(200, 140);
            this.Controls.Add(binSealLabel);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(700, 228);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            BinSealGrade.Validating += BinSealGrade_Validating;
            this.Controls.Add(this.BinSealGrade);

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.TextAlign = ContentAlignment.MiddleRight;
            firmnessLabel.Location = new System.Drawing.Point(500, 393);
            firmnessLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(firmnessLabel);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Location = new System.Drawing.Point(620, 378);
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
            this.Controls.Add(this.FirmnessBox);
            //
            ////


            //Delvicid Control
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.TextAlign = ContentAlignment.MiddleRight;
            delvicidLabel.Location = new System.Drawing.Point(500, 463);
            delvicidLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(delvicidLabel);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Location = new System.Drawing.Point(620, 448);
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
            this.Controls.Add(this.DelvicidBox);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Location = new System.Drawing.Point(580, 578);
            initialsLabel.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            this.Initials.Location = new System.Drawing.Point(660, 578);
            this.Initials.Size = new System.Drawing.Size(40, 30);
            Initials.Validating += Initials_Validating;
            this.Controls.Add(this.Initials);
        }

        private void InitializeScrap()
        {
            this.dateLabel = new Label();
            dateLabel.Text = "Lot Date:";
            dateLabel.TextAlign = ContentAlignment.MiddleRight;
            dateLabel.Location = new System.Drawing.Point(140, 218);
            this.Controls.Add(dateLabel);
            this.Date = new DateTimePicker();
            this.Date.Location = new System.Drawing.Point(250, 218);
            this.Date.Name = "Date Picker";
            this.Date.CustomFormat = "MM-dd-yyyy";
            this.Date.Format = DateTimePickerFormat.Custom;
            this.Date.Size = new System.Drawing.Size(140, 50);
            Date.Text = "01/01/2024";
            Date.Validating += LotDate_Validating;
            this.Controls.Add(this.Date);

            this.skidNumberLabel = new Label();
            skidNumberLabel.Text = "Skid/Tote Number:";
            skidNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            skidNumberLabel.Location = new System.Drawing.Point(100, 288);
            skidNumberLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(skidNumberLabel);
            this.ToteSkidNumber = new NumericUpDown();
            this.ToteSkidNumber.Location = new System.Drawing.Point(250, 298);
            this.ToteSkidNumber.Name = "Skid Number";
            this.ToteSkidNumber.Size = new System.Drawing.Size(70, 50);
            ToteSkidNumber.Text = "";
            ToteSkidNumber.Validating += ToteSkidNumber_Validating;
            this.Controls.Add(this.ToteSkidNumber);

            this.binWeightLabel = new Label();
            binWeightLabel.Text = "Bin Weight (lbs.):";
            binWeightLabel.TextAlign = ContentAlignment.MiddleRight;
            binWeightLabel.Location = new System.Drawing.Point(100, 448);
            binWeightLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(binWeightLabel);
            this.BinWeight = new NumericUpDown();
            this.BinWeight.Location = new System.Drawing.Point(250, 458);
            this.BinWeight.Name = "Bin Weight";
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 1200.00M;
            this.BinWeight.Value = 0.00M;
            BinWeight.Text = "";
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            BinWeight.Validating += BinWeight_Validating;
            this.Controls.Add(this.BinWeight);

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.TextAlign = ContentAlignment.MiddleRight;
            startTimeLabel.Location = new System.Drawing.Point(150, 528);
            startTimeLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 538);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.tempLabel = new Label();
            tempLabel.Text = "Temperature (F�):";
            tempLabel.TextAlign = ContentAlignment.MiddleRight;
            tempLabel.Location = new System.Drawing.Point(120, 608);
            tempLabel.Size = new System.Drawing.Size(120, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(250, 618);
            this.Temp.Name = "Temperature";
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = 0.00M;
            this.Temp.Maximum = 50.00M;
            this.Temp.Value = 0.00M;
            Temp.Text = "";
            this.Temp.Size = new System.Drawing.Size(100, 50);
            Temp.Validating += Temp_Validating;
            this.Controls.Add(this.Temp);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 748);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedScrap(sender, e); };

            this.binSealLabel = new Label();
            binSealLabel.Text = "       Bin Seal:                           (By checking this box you confirm that the bin is sealed adequately)";
            binSealLabel.TextAlign = ContentAlignment.MiddleRight;
            binSealLabel.Location = new System.Drawing.Point(460, 188);
            binSealLabel.Size = new System.Drawing.Size(200, 140);
            this.Controls.Add(binSealLabel);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(700, 228);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            BinSealGrade.Validating += BinSealGrade_Validating;
            this.Controls.Add(this.BinSealGrade);

            //Firmness Control
            //
            this.firmnessLabel = new Label();
            firmnessLabel.Text = "Firmness:";
            firmnessLabel.TextAlign = ContentAlignment.MiddleRight;
            firmnessLabel.Location = new System.Drawing.Point(500, 393);
            firmnessLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(firmnessLabel);
            this.FirmnessBox = new GroupBox();
            this.FirmnessBox.Name = "Firmness Box";
            this.FirmnessBox.Location = new System.Drawing.Point(620, 378);
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
            this.Controls.Add(this.FirmnessBox);
            //
            ////


            //Delvicid Control
            //
            //
            this.delvicidLabel = new Label();
            delvicidLabel.Text = "Delvicid Present:";
            delvicidLabel.TextAlign = ContentAlignment.MiddleRight;
            delvicidLabel.Location = new System.Drawing.Point(500, 463);
            delvicidLabel.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(delvicidLabel);
            this.DelvicidBox = new GroupBox();
            this.DelvicidBox.Name = "Firmness Box";
            this.DelvicidBox.Location = new System.Drawing.Point(620, 448);
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
            this.Controls.Add(this.DelvicidBox);
            //
            ////

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Location = new System.Drawing.Point(580, 578);
            initialsLabel.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            this.Initials.Location = new System.Drawing.Point(660, 578);
            this.Initials.Size = new System.Drawing.Size(40, 30);
            Initials.Validating += Initials_Validating;
            this.Controls.Add(this.Initials);
        }

        private void InitializePowder()
        {
            this.bagCountLabel = new Label();
            bagCountLabel.Text = "Bag Count:";
            bagCountLabel.TextAlign = ContentAlignment.MiddleRight;
            bagCountLabel.Location = new System.Drawing.Point(100, 288);
            bagCountLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(bagCountLabel);
            this.BagCount = new NumericUpDown();
            BagCount.Location = new System.Drawing.Point(250, 298);
            BagCount.Name = "Bag Count";
            BagCount.Size = new System.Drawing.Size(70, 50);
            BagCount.Maximum = 10;
            BagCount.Minimum = 0;
            BagCount.Increment = 1;
            BagCount.Value = 0;
            BagCount.Text = "";
            BagCount.Validating += BagCount_Validating;
            this.Controls.Add(this.BagCount);

            this.startTimeLabel = new Label();
            startTimeLabel.Text = "Start Time:";
            startTimeLabel.TextAlign = ContentAlignment.MiddleRight;
            startTimeLabel.Location = new System.Drawing.Point(150, 378);
            startTimeLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(startTimeLabel);
            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(250, 388);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.initialsLabel = new Label();
            initialsLabel.Text = "Initials:";
            initialsLabel.Location = new System.Drawing.Point(580, 478);
            initialsLabel.Size = new System.Drawing.Size(80, 30);
            initialsLabel.TextAlign = ContentAlignment.MiddleRight;
            this.Controls.Add(initialsLabel);
            this.Initials = new TextBox();
            this.Initials.Name = "Initials";
            this.Initials.Location = new System.Drawing.Point(670, 478);
            this.Initials.Size = new System.Drawing.Size(40, 30);
            Initials.Validating += Initials_Validating;
            this.Controls.Add(this.Initials);

            this.powderLotNumberLabel = new Label();
            powderLotNumberLabel.Text = "Lot Number:";
            powderLotNumberLabel.Location = new System.Drawing.Point(460, 388);
            powderLotNumberLabel.Size = new System.Drawing.Size(90, 60);
            powderLotNumberLabel.TextAlign = ContentAlignment.TopRight;
            this.Controls.Add(powderLotNumberLabel);
            this.PowderLotNumber = new TextBox();
            PowderLotNumber.Name = "Powder Lot Number";
            PowderLotNumber.Location = new System.Drawing.Point(560, 393);
            PowderLotNumber.Size = new System.Drawing.Size(90, 30);
            PowderLotNumber.Validating += PowderLotNumber_Validating;
            this.Controls.Add(PowderLotNumber);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 648);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);
            this.SubmitButton.Click +=
                delegate (object sender, EventArgs e) { SubmitButton_ClickedPowder(sender, e); };
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
                MessageBox.Show("Block was successfully tracked");
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
                    new { Column1 = Temp.Value.ToString() + "�F" , Column2 = "GOOD SEAL"},
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

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedTypeB(object sender, EventArgs e, string productNumber)
        {
            var data = new[]
                {
                    new { Column1 = productNumber, Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "�F" , Column2 = "GOOD SEAL"},
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

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedTypeC(object sender, EventArgs e, string productNumber)
        {
            var data = new[]
                {
                    new { Column1 = productNumber, Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "�F" , Column2 = "GOOD SEAL"},
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

            NewSelection();
            this.wb.Save();
        }

        private void SubmitButton_ClickedScrap(object sender, EventArgs e)
        {
            var data = new[]
                {
                    new { Column1 = "Scrap Tote", Column2 = this.Date.Value.ToShortDateString() },
                    new { Column1 = BinWeight.Value.ToString() + " lbs."  , Column2 = StartTime.Value.ToLongTimeString()  },
                    new { Column1 = Temp.Value.ToString() + "�F" , Column2 = "GOOD SEAL"},
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

            NewSelection();
            this.wb.Save();
        }


        //      VALIDATION METHODS

        private void LotDate_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            DateTime sixMonthsAgo = DateTime.Now.AddMonths(-6);
            DateTime sixMonthsFurther = DateTime.Now.AddMonths(6);

            if (this.Date.Value <= sixMonthsAgo || this.Date.Value >= sixMonthsFurther)
            {
                MessageBox.Show($"The lot date ({this.Date.Value.Date}) is not a valid lot date because it is six months or more away from today's date. Please enter a valid lot date");
                e.Cancel = true;
                errorProvider.SetError(Date, "Please Enter a Valid Date");
            }
        }

        private void ToteSkidNumber_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if(this.ToteSkidNumber.Value == 0 || this.ToteSkidNumber.Text == "")
            {
                MessageBox.Show("Please enter a valid skid/tote number. Number cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(ToteSkidNumber, "Please Enter a Valid Tote/Skid Number");
            }
        }

        private void BagCount_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if (this.BagCount.Value == 0 || this.BagCount.Text == "")
            {
                MessageBox.Show("Please enter a valid bag count. Number cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(BagCount, "Please Enter a Valid Bag Count");
            }
        }

        private void PowderLotNumber_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if(this.PowderLotNumber.Text.Length < 1 || this.PowderLotNumber.Text.Length > 14)
            {
                MessageBox.Show("Please enter a valid powder lot number.");
                e.Cancel = true;
                errorProvider.SetError(PowderLotNumber, "Please Enter a Valid Lot");
            }
            else if(Regex.IsMatch(PowderLotNumber.Text, @"[^a-zA-Z0-9]"))
            {
                MessageBox.Show("Lot number cannot contain special characters.");
                e.Cancel = true;
                errorProvider.SetError(PowderLotNumber, "Lot number cannot contain special characters.");
            }
            else
            {
                e.Cancel = false;
            }
        }

        private void BinWeight_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if (this.BinWeight.Value == 0.00M || this.BinWeight.Text == "")
            {
                MessageBox.Show("Please enter a valid weight. Weight cannot be 0.");
                e.Cancel = true;
                errorProvider.SetError(BinWeight, "Please Enter a Valid Weight");
            }
        }

        private void Temp_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if (Temp.Value == 0.00M || Temp.Text == "")
            {
                MessageBox.Show("Please enter a valid temperature.");
                e.Cancel = true;
                errorProvider.SetError(Temp, "Please Enter a Valid Temperature");
            }
        }

        private void BinSealGrade_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if(BinSealGrade.Checked == false)
            {
                MessageBox.Show("Please mark that you have verified the quality of the block's seal.");
                e.Cancel = true;
                errorProvider.SetError(BinSealGrade, "Please Check to Verify Seal Quality");
            }
        }

        private void Initials_Validating(object? sender, CancelEventArgs e)
        {
            errorProvider = new ErrorProvider();

            if(Initials.Text.Length < 2 || Initials.Text.Length > 3)
            {
                MessageBox.Show("Initials must be 2-3 letters in length.");
                e.Cancel = true;
                errorProvider.SetError(Initials, "Initials must be 2-3 letters in length");
            }
            else if(Regex.IsMatch(Initials.Text, @"\d") || Regex.IsMatch(Initials.Text, @"[^a-zA-Z0-9]"))
            {
                MessageBox.Show("Initials cannot contain numbers or special characters.");
                e.Cancel = true;
                errorProvider.SetError(Initials, "Initials cannot contain numbers or special characters.");
            }
            else
            {
                e.Cancel = false;
            }
        }
    }
}
