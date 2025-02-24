using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace shred_usage_writer
{
    public partial class ItemNumberControl : UserControl
    {
        public MaskedTextBox mtbFirstPart;
        public MaskedTextBox mtbSecondPart;
        public string ItemNumber => $"{mtbFirstPart.Text}-{mtbSecondPart.Text}";

        public ItemNumberControl()
        {
            InitializeComponent();
            InitializeControl();
        }

        private void InitializeControl()
        {
            // First part of item number (3 digits)
            mtbFirstPart = new MaskedTextBox("000") { Width = 50, TextAlign = HorizontalAlignment.Center };
            mtbFirstPart.TextChanged += ValidateInput;

            // Hyphen separator
            Label lblHyphen = new Label { Text = "-", AutoSize = true };

            // Second part of item number (6 digits)
            mtbSecondPart = new MaskedTextBox("000000") { Width = 80, TextAlign = HorizontalAlignment.Center };
            mtbSecondPart.TextChanged += ValidateInput;

            // Layout
            FlowLayoutPanel panel = new FlowLayoutPanel { AutoSize = true };
            panel.Controls.Add(mtbFirstPart);
            panel.Controls.Add(lblHyphen);
            panel.Controls.Add(mtbSecondPart);
            Controls.Add(panel);
        }

        private void ValidateInput(object sender, EventArgs e)
        {
            bool valid = mtbFirstPart.Text.Length == 3 && mtbSecondPart.Text.Length == 6;
            this.BackColor = valid ? SystemColors.Control : System.Drawing.Color.LightCoral;
        }
    }
}
