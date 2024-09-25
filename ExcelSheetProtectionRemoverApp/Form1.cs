using ProtectionRemoverLib;
using System;
using System.Security;
using System.Windows.Forms;

namespace ExcelSheetProtectionRemoverApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "Excel取消工作表保護";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    foreach (var file in openFileDialog1.FileNames)
                    {
                        ProtectionRemover.RemoveProtection(file);
                    }
                    MessageBox.Show($"OK");

                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
    }
}
