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
        internal System.Windows.Forms.RadioButton Firmness;
        internal System.Windows.Forms.RadioButton Delvicid;
        internal System.Windows.Forms.TextBox Initials;
        //Bin Seal Grade
        //Firmness (F or S)
        //Delvicid Y or N
        //Initials


        // Initialize ComboBox1.
        private void InitializeComboBox()
        {
            this.ComboBox1 = new ComboBox();
            this.ComboBox1.Location = new System.Drawing.Point(100, 38);
            this.ComboBox1.Name = "ComboBox1";
            this.ComboBox1.Size = new System.Drawing.Size(200, 50);
            this.ComboBox1.TabIndex = 0;
            this.ComboBox1.Text = "SELECT PRODUCT";
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
            string selectedProduct = (string) comboBox.SelectedItem;

            switch(selectedProduct)
            {
                case "001-000133":
                    this.ws = wb.Worksheet("1-133");
                    this.ws.Worksheet.Cell("A67").Value = "Hello World";
                    this.wb.Save();
                    break;
                case "001-000169":
                    NewSelection();
                    InitializeBlockTypeA();
                    break;
                case "001-000195":
                    NewSelection();
                    break;
                case "001-000229":
                    break;
                case "001-000360":
                    break;
                case "001-000455":
                    break;
                case "001-000470":
                    break;
                case "001-000508":
                    break;
                case "001-000525":
                    break;
                case "001-000528":
                    break;
                case "008-000005 PS Purchased":
                    break;
                case "008-000021 WM Purchased":
                    break;
                case "008-000001 Asiago":
                    break;
                case "008-000002 Cheddar":
                    break;
                case "008-000006 Parmesan":
                    break;
                case "008-000010 White Ched":
                    break;
                case "008-000022 Meunster":
                    break;
                case "008-000007 Provolone":
                    break;
                case "002-000035 Scrap":
                    break;
                case "Powder":
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
            this.Date = new DateTimePicker();
            this.Date.Location = new System.Drawing.Point(150, 118);
            this.Date.Name = "Date Picker";
            this.Date.CustomFormat = "MM-dd-yyyy";
            this.Date.Format = DateTimePickerFormat.Custom;
            this.Date.Size = new System.Drawing.Size(140, 50);
            this.Controls.Add(this.Date);

            this.ToteSkidNumber = new NumericUpDown();
            this.ToteSkidNumber.Location = new System.Drawing.Point(150, 198);
            this.ToteSkidNumber.Name = "Skid Number";
            this.ToteSkidNumber.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.ToteSkidNumber);

            this.NumberPieces = new NumericUpDown();
            this.NumberPieces.Location = new System.Drawing.Point(150, 278);
            this.NumberPieces.Name = "Number of Pieces";
            this.NumberPieces.Minimum = 0;
            this.NumberPieces.Maximum = 160;
            this.NumberPieces.Value = 160;
            this.NumberPieces.Size = new System.Drawing.Size(70, 50);
            this.Controls.Add(this.NumberPieces);

            this.BinWeight = new NumericUpDown();
            this.BinWeight.Location = new System.Drawing.Point(150, 358);
            this.BinWeight.Name = "Bin Weight";
            this.BinWeight.DecimalPlaces = 2;
            this.BinWeight.Increment = 0.01M;
            this.BinWeight.Minimum = 0.00M;
            this.BinWeight.Maximum = 980.00M;
            this.BinWeight.Value = 960.00M;
            this.BinWeight.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.BinWeight);

            this.StartTime = new DateTimePicker();
            this.StartTime.CustomFormat = "hh':'mm";
            this.StartTime.Format = DateTimePickerFormat.Custom;
            this.StartTime.ShowUpDown = true;
            this.StartTime.Location = new System.Drawing.Point(150, 438);
            this.StartTime.Name = "Start Time";
            this.StartTime.Size = new System.Drawing.Size(100, 50);
            this.Controls.Add(this.StartTime);

            this.Temp = new NumericUpDown();
            this.Temp.Location = new System.Drawing.Point(150, 518);
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
            this.SubmitButton.Location = new System.Drawing.Point(170, 598);
            this.SubmitButton.Size = new System.Drawing.Size(110, 40);
            this.SubmitButton.Text = "SUBMIT";
            this.Controls.Add(this.SubmitButton);

            this.BinSealGrade = new CheckBox();
            this.BinSealGrade.Name = "Bin Seal Grade";
            this.BinSealGrade.Location = new System.Drawing.Point(500, 118);
            this.BinSealGrade.Size = new System.Drawing.Size(20, 20);
            this.Controls.Add(this.BinSealGrade);

            this.Firmness = new RadioButton();
            this.Firmness.Name = "Firmness";
            this.Firmness.Location = new System.Drawing.Point(450, 198);
            this.Firmness.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(this.Firmness);

            this.Delvicid = new RadioButton();
            this.Delvicid.Name = "Delvicid";
            this.Delvicid.Location = new System.Drawing.Point(450, 278);
            this.Delvicid.Size = new System.Drawing.Size(80, 30);
            this.Controls.Add(this.Delvicid);
        }
    }
}
