using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shred_usage_writer
{

    public class CustomMessageBox : Form
    {
        private DataGridView dataGridView;
        private Button yesButton;
        private Button noButton;
        private Panel buttonPanel;
        public DialogResult Result { get; private set; }

        public CustomMessageBox()
        {
            this.Size = new Size(450, 400);

            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                Width = 400
            };

            buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70
            };

            yesButton = new Button
            {
                Text = "YES",
                DialogResult = DialogResult.Yes,
                Dock = DockStyle.Left,
                Width = this.ClientSize.Width/2
            };

            noButton = new Button
            {
                Text = "NO",
                DialogResult = DialogResult.No,
                Dock = DockStyle.Right,
                Width = this.ClientSize.Width/2
            };

            

            // Add buttons to the panel
            buttonPanel.Controls.Add(yesButton);
            buttonPanel.Controls.Add(noButton);

            yesButton.Click += (sender, e) => { Result = DialogResult.Yes; Close(); };
            noButton.Click += (sender, e) => { Result = DialogResult.No; Close(); };

            Controls.Add(dataGridView);
            Controls.Add(buttonPanel);
        }

        public void SetData(object dataSource)
        {
            dataGridView.DataSource = dataSource;
        }

        public static DialogResult Show(object dataSource)
        {
            using (var form = new CustomMessageBox())
            {
                form.SetData(dataSource);
                form.ShowDialog();
                return form.Result;
            }
        }
    }

}
