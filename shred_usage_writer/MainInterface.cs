using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace shred_usage_writer
{
    public partial class MainInterface : Form
    {
        public XLWorkbook wb;
        public IXLWorksheet ws;
        public MainInterface()
        {
            InitializeComponent();
            this.Text = "Miceli Dairy Products - Block Usage Reporting Tool";
            InitializeComboBox();
            this.wb = new XLWorkbook(@"C:\Users\pbsmi\workspace\Excel-Sharp\test_shred_usage.xlsx");
        }

        // Declare Components
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.DateTimePicker Date;
        internal System.Windows.Forms.NumericUpDown ToteSkidNumber;
        internal System.Windows.Forms.NumericUpDown NumberPieces;
        internal System.Windows.Forms.NumericUpDown BinWeight;
        internal System.Windows.Forms.DateTimePicker StartTime;
        internal System.Windows.Forms.NumericUpDown Temp;
        internal System.Windows.Forms.Button SubmitButton;
        internal System.Windows.Forms.CheckBox BinSealGrade;
        internal System.Windows.Forms.GroupBox FirmnessBox;
        internal System.Windows.Forms.RadioButton FirmnessFirm;
        internal System.Windows.Forms.RadioButton FirmnessSoft;
        internal System.Windows.Forms.GroupBox DelvicidBox;
        internal System.Windows.Forms.RadioButton DelvicidTrue;
        internal System.Windows.Forms.RadioButton DelvicidFalse;
        internal System.Windows.Forms.TextBox Initials;


        // Initialize ComboBox1.
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
                    this.ws = wb.Worksheet("1-133");
                    this.ws.Worksheet.Cell("A67").Value = "Hello World";
                    this.wb.Save();
                    break;
                case "001-000169":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000169      Part-Skim Block";
                    InitializeBlockTypeA();
                    break;
                case "001-000195":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000195      ";
                    break;
                case "001-000229":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000229      ";
                    break;
                case "001-000360":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000360      ";
                    break;
                case "001-000455":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000455      ";
                    break;
                case "001-000470":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000470      ";
                    break;
                case "001-000508":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000508      ";
                    break;
                case "001-000525":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000525      Whole Milk Block";
                    break;
                case "001-000528":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 001-000528      ";
                    break;
                case "008-000005 PS Purchased":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000005      Purchased PS Block";
                    break;
                case "008-000021 WM Purchased":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000021      Purchased WM Block";
                    break;
                case "008-000001 Asiago":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000001      ";
                    break;
                case "008-000002 Cheddar":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000002      ";
                    break;
                case "008-000006 Parmesan":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000006      ";
                    break;
                case "008-000010 White Ched":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000010      ";
                    break;
                case "008-000022 Meunster":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000022      ";
                    break;
                case "008-000007 Provolone":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 008-000007      ";
                    break;
                case "002-000035 Scrap":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 002-000035      ";
                    break;
                case "Powder":
                    NewSelection();
                    this.Text = "Miceli Dairy Products - Block Usage Reporting Tool                 Powder      ";
                    break;
                default:
                    MessageBox.Show("Please Make a Product Selection");
                    break;

            }
        }

        private void NewSelection()
        {
            foreach (System.Windows.Forms.Control item in this.Controls)
            {
                if (item != ComboBox1)
                {
                    this.Controls.Remove(item);
                }
            }

            //Removing Numeric UpDowns
            this.Controls.Remove(ToteSkidNumber);
            this.Controls.Remove(BinWeight);
            this.Controls.Remove(Temp);
        }

        private void InitializeBlockTypeA()
        {
            Label dateLabel = new Label();
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
            this.Controls.Add(this.Date);

            Label skidNumberLabel = new Label();
            skidNumberLabel.Text = "Skid/Tote Number:";
            skidNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            skidNumberLabel.Location = new System.Drawing.Point(100, 288);
            skidNumberLabel.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(skidNumberLabel);
            this.ToteSkidNumber = new NumericUpDown();
            this.ToteSkidNumber.Location = new System.Drawing.Point(250, 298);
            this.ToteSkidNumber.Name = "Skid Number";
            this.ToteSkidNumber.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.ToteSkidNumber);

            Label piecesNumberLabel = new Label();
            piecesNumberLabel.Text = "Number of Pieces:";
            piecesNumberLabel.TextAlign = ContentAlignment.MiddleRight;
            piecesNumberLabel.Location = new System.Drawing.Point(145, 368);
            piecesNumberLabel.Size = new System.Drawing.Size(90, 50);
            this.Controls.Add(piecesNumberLabel);
            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Location = new System.Drawing.Point(250, 378);
            this.NumberPieces.Name = "Number of Pieces";
            this.NumberPieces.Minimum = 0;
            this.NumberPieces.Maximum = 160;
            this.NumberPieces.Value = 160;
            this.NumberPieces.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.NumberPieces);

            Label binWeightLabel = new Label();
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
            this.BinWeight.Maximum = 980.00M;
            this.BinWeight.Value = 960.00M;
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.BinWeight);

            Label startTimeLabel = new Label();
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

            Label tempLabel = new Label();
            tempLabel.Text = "Temperature (F°):";
            tempLabel.TextAlign = ContentAlignment.MiddleRight;
            tempLabel.Location = new System.Drawing.Point(120, 608);
            tempLabel.Size = new System.Drawing.Size(120, 50);
            this.Controls.Add(tempLabel);
            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(250, 618);
            this.Temp.Name = "Temperature";
            this.Temp.DecimalPlaces = 2;
            this.Temp.Increment = 0.1M;
            this.Temp.Minimum = -20.00M;
            this.Temp.Maximum = 70.00M;
            this.Temp.Value = 32.00M;
            this.Temp.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.Temp);

            this.SubmitButton = new Button();
            this.SubmitButton.Name = "Submit";
            this.SubmitButton.Location = new System.Drawing.Point(470, 748);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);

            Label binSealLabel = new Label();
            binSealLabel.Text = "       Bin Seal:               (By checking this box you confirm that the bin is sealed adequately)";
            binSealLabel.TextAlign = ContentAlignment.MiddleRight;
            binSealLabel.Location = new System.Drawing.Point(460, 188);
            binSealLabel.Size = new System.Drawing.Size(200, 140);
            this.Controls.Add(binSealLabel);
            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(700, 228);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            this.Controls.Add(this.BinSealGrade);

            //Firmness Control
            //
            Label firmnessLabel = new Label();
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
            Label delvicidLabel = new Label();
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
            trueDelvicidLabel.Size = new System.Drawing.Size(60, 30);
            this.DelvicidBox.Controls.Add(trueDelvicidLabel);
            this.DelvicidTrue = new RadioButton();
            this.DelvicidTrue.Name = "Delvicid True";
            this.DelvicidTrue.Location = new System.Drawing.Point(80, 25);
            this.DelvicidTrue.Size = new System.Drawing.Size(40, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidTrue);
            //  Second Button
            Label falseDelvicidLabel = new Label();
            falseDelvicidLabel.Text = "No";
            falseDelvicidLabel.Location = new System.Drawing.Point(110, 25);
            falseDelvicidLabel.Size = new System.Drawing.Size(60, 30);
            this.FirmnessBox.Controls.Add(falseDelvicidLabel);
            this.DelvicidFalse = new RadioButton();
            this.DelvicidFalse.Name = "Delvicid False";
            this.DelvicidFalse.Location = new System.Drawing.Point(170, 25);
            this.DelvicidFalse.Size = new System.Drawing.Size(30, 30);
            this.DelvicidBox.Controls.Add(this.DelvicidFalse);
            this.Controls.Add(this.DelvicidBox);
            //
            ////
        }

        
    }
}
