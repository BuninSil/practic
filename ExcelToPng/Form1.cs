using System;
using System.IO;
using System.Windows.Forms;
//
using OfficeOpenXml;
//
using PdfSharp.Pdf;
using PdfSharp.Drawing;
//
using VkNet;
using VkNet.Model;

namespace ExcelToPdf
{
    public partial class Form1 : Form
    {
        private string selectedFilePath;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = openFileDialog.FileName;
                //MessageBox.Show("File selected: " + selectedFilePath);
                label1.Text = "Выбрано: " + selectedFilePath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Выберите сначала файл");
                return;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(selectedFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Arial", 12);

                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        string cellText = worksheet.Cells[row, col].Text;
                        gfx.DrawString(cellText, font, XBrushes.Black, col * 50, row * 20);
                    }
                }

                string savePath = Path.ChangeExtension(selectedFilePath, "pdf");

                using (MemoryStream stream = new MemoryStream())
                {
                    pdf.Save(stream, false);
                    File.WriteAllBytes(savePath, stream.ToArray());
                }

                MessageBox.Show("PDF файл сохранен: " + savePath);
                var Api = new VkApi();
                Api.Authorize(new ApiAuthParams
                {
                    AccessToken = "vk1.a.dkYMWiXHBExYArXHSE9d44E0OYawqcn25IIHCL5tEKAAL49gi4OvcHj6gddoWCPMI4MuZ37l53j13Dd-i2LCoiC7SQs0EptVlnJ9nFtXnB9AwLF6dGD2gT3Lw2Y6X3RyxJsANewJe2wQYIWSJAEw0hsdGXlLW1DoEmXWSXGHxLTVCJLQLvPWcNuEIuQTKRaWcnzSvbIC8L34H2nEstDcEg"
                });
                long userId = 571881967;
                Random rnd = new Random();
                int rundom = rnd.Next(0, 9999);
                Api.Messages.Send(new MessagesSendParams
                {
                    RandomId = rundom,
                    UserId = userId,
                    Message = $"Сконвертировано {savePath}",

                });
            }

            


        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }


    }
}
