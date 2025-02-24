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
        private ErrorProvider errorProvider = new ErrorProvider();

        public string ItemNumber => $"{mtbFirstPart.Text}-{mtbSecondPart.Text}";

        public ItemNumberControl()
        {
            InitializeComponent();
            InitializeControl();
            this.CausesValidation = true; // Ensure validation is enabled
        }

        private void InitializeControl()
        {
            // First part of item number (3 digits)
            mtbFirstPart = new MaskedTextBox("000")
            {
                Width = 60,
                TextAlign = HorizontalAlignment.Center,
                Font = new System.Drawing.Font("Arial", 14),
                Size = new System.Drawing.Size(60, 70),
                TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
            };

            // Hyphen separator
            Label lblHyphen = new Label
            {
                Text = "-",
                AutoSize = true,
                Font = new System.Drawing.Font("Arial", 16)
            };

            // Second part of item number (6 digits)
            mtbSecondPart = new MaskedTextBox("000000")
            {
                Width = 120,
                TextAlign = HorizontalAlignment.Center,
                Font = new System.Drawing.Font("Arial", 14),
                Size = new System.Drawing.Size(120, 70),
                TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
            };

            // Layout
            FlowLayoutPanel panel = new FlowLayoutPanel { AutoSize = true };
            panel.Controls.Add(mtbFirstPart);
            panel.Controls.Add(lblHyphen);
            panel.Controls.Add(mtbSecondPart);
            Controls.Add(panel);
        }

        // Override OnValidating for automatic validation
        protected override void OnValidating(CancelEventArgs e)
        {
            base.OnValidating(e); // Call base class validation

            string firstPart = mtbFirstPart.Text.Trim();
            string secondPart = mtbSecondPart.Text.Trim();

            if (firstPart.Length != 3 || secondPart.Length != 6)
            {
                e.Cancel = true;
                this.BackColor = System.Drawing.Color.LightCoral;
                errorProvider.SetError(this, "Invalid Item Number. Format: XXX-XXXXXX (e.g., 123-456789)");
            }
            else
            {
                this.BackColor = SystemColors.Control;
                errorProvider.SetError(this, ""); // Clear error
            }
        }
    }

}
