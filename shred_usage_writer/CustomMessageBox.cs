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
        private Panel statementPanel;
        private TextBox statement;
        private Button yesButton;
        private Button noButton;
        private Panel buttonPanel;
        public DialogResult Result { get; private set; }

        public CustomMessageBox()
        {
            this.Size = new Size(450, 400);
            this.StartPosition = FormStartPosition.CenterParent;

            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                Width = 450,
                ColumnHeadersVisible = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
            };

            statementPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100
            };

            buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70
            };

            statement = new TextBox
            {
                Text = "Please Review and Confirm Before Submitting",
                Width = 450,
                TextAlign = HorizontalAlignment.Center
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
            statementPanel.Controls.Add( statement );
            buttonPanel.Controls.Add(yesButton);
            buttonPanel.Controls.Add(noButton);

            yesButton.Click += (sender, e) => { Result = DialogResult.Yes; Close(); };
            noButton.Click += (sender, e) => { Result = DialogResult.No; Close(); };

            Controls.Add(dataGridView);
            Controls.Add(buttonPanel);
            Controls.Add(statementPanel);
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
