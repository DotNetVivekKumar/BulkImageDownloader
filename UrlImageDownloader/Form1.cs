using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UrlImageDownloader
{
    public partial class Form1 : Form
    {
        string xsltFile = Application.StartupPath + @"\\SampleFile.xlsx";
        string imagePath = Application.StartupPath + "/Images/";

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> readfile = ReadExcelFile();
            bool val = DownloadFile(readfile);
        }

        private List<string> ReadExcelFile()
        {
            List<string> imageUrls = new List<string>();

            try
            {
                // Open Excel file using EPPlus library
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(xsltFile)))
                {
                    // Get the first worksheet from the workbook
                    var myWorksheet = xlPackage.Workbook.Worksheets.First();

                    // Get the number of rows and columns in the worksheet
                    int rows = myWorksheet.Dimension.Rows;
                    int columns = myWorksheet.Dimension.Columns;

                    // Iterate through each row starting from the second row
                    for (int i = 2; i <= rows; i++)
                    {
                        // Get the content from the second column of the current row
                        string content = Convert.ToString(myWorksheet.Cells[i, 2].Value);

                        // Check if the content is not null or empty, then add it to the list
                        if (!string.IsNullOrEmpty(content))
                            imageUrls.Add(content);
                    }
                }
            }
            catch (Exception ex)
            {
                // Throw the exception to be handled by the caller
                throw ex;
            }

            // Return the list of image URLs
            return imageUrls;
        }


        private System.Drawing.Image DownloadImageFromUrl(string imageUrl)
        {
            System.Drawing.Image image = null;

            try
            {
                // Create a HttpWebRequest to download the image from the provided URL
                System.Net.HttpWebRequest webRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(imageUrl);

                // Set properties of the web request
                webRequest.AllowWriteStreamBuffering = false; // Disable buffering for large files
                webRequest.Timeout = 1200 * 1000; // Set timeout to 1200 seconds (20 minutes)

                // Get the response from the web request
                System.Net.WebResponse webResponse = webRequest.GetResponse();

                // Get the response stream
                System.IO.Stream stream = webResponse.GetResponseStream();

                // Create an image object from the response stream
                image = System.Drawing.Image.FromStream(stream);

                // Close the web response
                webResponse.Close();
            }
            catch (Exception ex)
            {
                // If an exception occurs, return null
                return null;
            }

            // Return the downloaded image
            return image;
        }


        protected bool DownloadFile(List<string> readfile)
        {
            try
            {
                // Loop through each image URL in the list
                foreach (string imgurl in readfile)
                {
                    // Get the filename from the URL
                    string fileN = Path.GetFileName(imgurl);

                    // Download the image from the URL
                    System.Drawing.Image image = DownloadImageFromUrl(imgurl);

                    // Check if the image was successfully downloaded
                    if (image != null)
                    {
                        // Construct the file path to save the image
                        string fileName = System.IO.Path.Combine(imagePath, fileN);

                        // Save the image to the specified file path
                        image.Save(fileName);
                    }
                    else
                    {
                        // If image download fails, log the URL to a text file
                        StreamWriter writer = new StreamWriter(Application.StartupPath + "/log.txt", true);
                        writer.WriteLine("Url-- " + imgurl);
                        writer.Close();
                    }
                }
                // Return true if all images were downloaded successfully
                return true;
            }
            catch (Exception ex)
            {
                // If an exception occurs during the download process, return false
                return false;
            }
        }

    }
}
