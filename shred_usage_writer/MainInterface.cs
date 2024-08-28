using ClosedXML.Excel;
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
            this.wb = new XLWorkbook(@"C:\Users\psmith\Desktop\test_shred_usage.xlsx");
        }

        // Declare Components
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.TextBox ToteSkidNumber;
        internal System.Windows.Forms.TextBox NumberPieces;
        internal System.Windows.Forms.TextBox BinWeight;
        internal System.Windows.Forms.DateTimePicker StartTime;
        internal System.Windows.Forms.TextBox Temp;

        //Bin Seal Grade
        //Firmness (F or S)
        //Delvicid Y or N
        //Initials


        // Initialize ComboBox1.
        private void InitializeComboBox()
        {
            this.ComboBox1 = new ComboBox();
            this.ComboBox1.Location = new System.Drawing.Point(128, 48);
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

        // Handles the ComboBox1 DropDown event. If the user expands the  
        // drop-down box, a message box will appear, recommending the
        // typical installation.
        private void ComboBox1_DropDown(object sender, System.EventArgs e)
        {
            
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
                    this.ws = wb.Worksheet("1-169");
                    this.ws.Worksheet.Cell("A67").Value = "Hello World";
                    this.wb.Save();
                    break;
                case "001-000195":
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

    }
}
