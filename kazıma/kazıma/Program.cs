using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Opera;
using Spire.Doc;
using Spire.Doc.Documents;


namespace kazıma
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string url = ""; //url
            string operaDriverPath = @"..."; // OperaDriver konumu
            string operaPath = @"..."; // Opera tarayıcı yolu
            OperaOptions options = new OperaOptions();
            options.BinaryLocation = operaPath;
            // WebDriver başlat
            IWebDriver driver = new OperaDriver(operaDriverPath, options);

            //word
            Document dosya = new Document();
            
            for (int i = 0; i < 400; i++) //max 500 sayfa
            {

                // URL'ye git
                driver.Navigate().GoToUrl(url);

                // Elemanı bul
                var elements = driver.FindElements(By.CssSelector("p[data-original]")); //verinin olduğu element

                string allText = "";
                // data-original değerini al
                foreach (var element in elements)
                {
                    try
                    {
                        string dataOriginal = element.GetAttribute("data-original");
                        allText += dataOriginal+"\n";
                    }
                    catch { }

                }

                if (allText != null)
                {
                    Section section = dosya.AddSection();
                    Paragraph paragraph = section.AddParagraph();
                    paragraph.AppendText(allText);
                }

                // "Sonraki Bölüm" butonunu bul ve tıkla

                bool findbutton = false;
                while (!findbutton)
                {
                    try
                    {
                        var nextButton = driver.FindElement(By.CssSelector("a[rel='next']")); //buton ile sonraki sayfaya geçmek
                        url = nextButton.GetAttribute("href");
                        findbutton = true;
                    }
                    catch { Thread.Sleep(500); }
                }
            }

            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "file.docx");
            dosya.SaveToFile(filePath);
            Console.WriteLine($"Word belgesi kaydedildi: {filePath}");

        }
    }
}
