using System;
using System.Drawing;
using System.IO;
using System.Reflection.Emit;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using static System.Net.Mime.MediaTypeNames;

namespace GoogleSearchTest
{
    public class GoogleSearchTest
    {
        private IWebDriver driver;
        Document pdfDoc = new Document();
        private int counter = 1;
        public void SetUp()
        {
            // Membuat objek ChromeDriver
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
        }

        public void RegisterMobipad()
        {

            initiatePDF("Mobipad");
            // Membuka halaman Google
            
            driver.Navigate().GoToUrl("https://stage.mobipaid.com/en/register");

            // Menunggu hasil pencarian muncul

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            captureImagetoPDF(pdfDoc, "Halaman Bagian 1", "Halaman Registrasi");


            String firstName, LastName, dtEmail, dtPassword, dtCompany, dtNoHP, dtNegara, dtProvinsi;
            // Mencari sesuatu di Google
            firstName = "ilyas";
            LastName = "Dewanto";
            dtEmail = "Testemail@gmail.com";
            dtPassword = "Password123.";
            dtCompany = "ilyCorp";
            dtNoHP = "89887999876";
            


            IWebElement inputFirstName = driver.FindElement(By.Name("signatory_first_name"));
            IWebElement inputLastName = driver.FindElement(By.Name("signatory_last_name"));
            IWebElement inputEmail = driver.FindElement(By.Name("email"));
            IWebElement inputPassword = driver.FindElement(By.Name("password"));
            IWebElement inputCompanyName = driver.FindElement(By.Name("name"));
            
            IWebElement selectNegara = driver.FindElement(By.Id("country"));
            IWebElement inputProvinsi = driver.FindElement(By.Name("state"));
            IWebElement checkBoxSetuju = driver.FindElement(By.ClassName("psa-checkbox"));
            inputFirstName.SendKeys(firstName);
            inputLastName.SendKeys(LastName);
            inputEmail.SendKeys(dtEmail);
            inputPassword.SendKeys(dtPassword);
            inputCompanyName.SendKeys(dtCompany);
           
            
            inputProvinsi.SendKeys("Jawa Tengah");
            checkBoxSetuju.Click();
            captureImagetoPDF(pdfDoc, "Input Form", "Halaman Registrasi");





            // Menutup laporan PDF
            pdfDoc.Close();
        }

        public void TearDown()
        {
            // Menutup ChromeDriver
            driver.Close();
            driver.Quit();
        }


        public void captureImagetoPDF(Document pdfDoc, string caption, string label)
        {
            tempImage();
            AddImageToPDF(pdfDoc, @"D:\WWW\Testing\reeport\MyScreenshot.png", caption, label);
        }
        public void tempImage()
        {
            // Mengambil screenshot halaman hasil pencarian
            Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
            string screenshotFileName = @"D:\WWW\Testing\reeport\MyScreenshot.png";
            screenshot.SaveAsFile(screenshotFileName, ScreenshotImageFormat.Png);
        }
        public void AddImageToPDF(Document pdfDoc, string imagePath, string caption, string label)
        {
            // Menambahkan gambar ke dalam laporan PDF
            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(imagePath);
            image.ScaleToFit(pdfDoc.PageSize.Width - pdfDoc.LeftMargin - pdfDoc.RightMargin, pdfDoc.PageSize.Height - pdfDoc.TopMargin - pdfDoc.BottomMargin);

            // Menambahkan caption dan label ke dalam laporan PDF
            Paragraph captionPara = new Paragraph(caption, new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL));
            captionPara.SpacingBefore = 10f;
            captionPara.SpacingAfter = 20f; // Menambahkan margin bawah pada caption
            Paragraph labelPara = new Paragraph(counter.ToString() + ". " + label, new Font(Font.FontFamily.HELVETICA, 15, Font.BOLD));
            labelPara.SpacingBefore = 20f; // Menambahkan margin atas pada label
            labelPara.SpacingAfter = 10f;
            counter++;

            if (counter % 2 == 0) // Jika jumlah gambar sudah genap
            {
                pdfDoc.NewPage(); // Tambahkan halaman baru
            }

            // Menambahkan gambar, caption, dan label ke dalam laporan PDF
            pdfDoc.Add(labelPara);
            pdfDoc.Add(image);
            pdfDoc.Add(captionPara);
        }

        public void initiatePDF(string tittleName)
        {
            // Membuat laporan PDF
            string pdfFileName = @"D:\WWW\Testing\reeport\" + tittleName + ".pdf";
            PdfWriter.GetInstance(pdfDoc, new FileStream(pdfFileName, FileMode.Create));
            pdfDoc.Open();

            string subTitle = "Regression Test";
            string docName = "Automation Test Execution Document";

            // Menambahkan judul pada halaman pertama
            //Paragraph titlePara = new Paragraph(tittleName, new Font(Font.FontFamily.HELVETICA, 24, Font.BOLD)); COMMENT BENTAR
            Paragraph titlePara = new Paragraph(tittleName, new Font(Font.FontFamily.HELVETICA, 36, Font.BOLD));
            Paragraph subTitlePara = new Paragraph(subTitle, new Font(Font.FontFamily.HELVETICA, 20, Font.BOLD));
            Paragraph docNamePara = new Paragraph(docName, new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD));

            //titlePara.Alignment = Element.ALIGN_CENTER;   COMMENT BENTAR
            titlePara.Alignment = Element.ALIGN_RIGHT;
            subTitlePara.Alignment = Element.ALIGN_RIGHT;
            docNamePara.Alignment = Element.ALIGN_LEFT;

            titlePara.SpacingBefore = 100f; // menambahkan jarak antara judul dan konten berikutnya
            titlePara.SpacingAfter = 10f;  // menambahkan jarak antara judul dan konten berikutnya
            subTitlePara.SpacingAfter = 80f; // menambahkan jarak antara sub judul dan konten berikutnya
            docNamePara.SpacingAfter = 10f; // menambahkan jarak antara dokumen name dan konten berikutnya

            //iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(@"C:\Users\Windows\Downloads\WhatsApp Image 2023 - 01 - 05 at 15.40.14.jpeg");
            //image.ScaleToFit(pdfDoc.PageSize.Width - pdfDoc.LeftMargin - pdfDoc.RightMargin, pdfDoc.PageSize.Height - pdfDoc.TopMargin - pdfDoc.BottomMargin);

            pdfDoc.Add(titlePara);
            pdfDoc.Add(subTitlePara);
            pdfDoc.Add(docNamePara);
            pdfDoc.NewPage();
        }




        static void Main(string[] args)
        {
            GoogleSearchTest test = new GoogleSearchTest();
            test.SetUp();
            test.RegisterMobipad();
            test.TearDown();
        }
    }
}