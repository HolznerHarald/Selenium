using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

//using Microsoft.Office.Interop.Excel;

namespace Selenium
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainWindow MW;
        public MainWindow()
        {
            InitializeComponent();
            MW = System.Windows.Application.Current.MainWindow as MainWindow;
        }

        private void WerteAbfragen_Click(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.Hilfs();
            g9.AktWerteSchreiben();
        }
        private void AndersAbfragen_Click(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.Hilfs();
            g9.AndersWerteSchreiben();
        }

        private void Hilfs_Click(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.Hilfs();

        }

        // ************* ab hier nur Test-Routinen

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.test();
            g9.closeBrowser();
            // System.Windows.MessageBox.Show("Click here!");

        }
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.test("https://demo.guru99.com/test/guru99home/");

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.test1("https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008");

        }
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            C1();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            // Instanziieren Sie ein Workbook objekt, das eine Excel Datei darstellt.
            Workbook wb = new Workbook();
            // Wenn Sie eine neue Arbeitsmappe erstellen, wird der Arbeitsmappe standardmäßig „Sheet1“ hinzugefügt.
            Worksheet sheet = wb.Worksheets[0];
            // Greifen Sie auf die Zelle "A1" im Blatt zu.
            Cell cell = sheet.Cells["A1"];

            // Geben Sie das "Hello World!" Text in die Zelle "A1".
            cell.PutValue("Hello World!");

            // Speichern Sie die Excel Datei als .xlsx Datei.
            wb.Save(@"C:\C#\Excel11.xlsx", SaveFormat.Xlsx);
        }

        public void C1()
        {
            string strFilePath = @"C:\C#\Data.csv";
            string strSeperator = ",";
            StringBuilder sbOutput = new StringBuilder();

            int[][] inaOutput = new int[][] { new int[] { 1000, 2000, 3000, 4000, 5000 },
                                        new int[] { 6000, 7000, 8000, 9000, 10000 },
                                        new int[] { 11000, 12000, 13000, 14000, 15000 } };

            int ilength = inaOutput.GetLength(0);
            for (int ii = 0; ii < ilength; ii++)
                sbOutput.AppendLine(string.Join(strSeperator, inaOutput[ii]));

            // Create and write the csv file
            File.WriteAllText(strFilePath, sbOutput.ToString());

            // To append more lines to the csv file
            File.AppendAllText(strFilePath, sbOutput.ToString());
        }
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.test2("https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008");
        }
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            Guru99Demo g9 = new Guru99Demo(MW);
            g9.startBrowser();
            g9.test3("https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008");
        }
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            DateTime Schleifenzeit = DateTime.Now;

            string hs = Schleifenzeit.ToString();

            Ausgabe.Text = hs.Substring(hs.Length - 8) + "\n";
            Ausgabe.Text = hs.Substring(0, 8);
        }
        private void PopUpTest1(object sender, RoutedEventArgs e)
        {  // Von Würfelprgramm wird Alert von Anleitung automatisch gestartet und dann mit accept geschlossen


            IWebDriver driver = new ChromeDriver("C:\\Users\\holzn\\.cache\\selenium\\chromedriver\\win64");

            driver.Url = "C:/C%23/Wuerfel/Wuerfel/index.html";

            //This step produce an alert on screen, Anleitung als JS-Alert

            driver.FindElement(By.XPath("//*[@id='Anleitung']")).Click();
            

            Thread.Sleep(5000);

            IAlert simpleAlert = driver.SwitchTo().Alert();
            simpleAlert.Accept();

        }
        private void PopUpTest2(object sender, RoutedEventArgs e)
        {  // Boerse.de starten, dann das iframe Fenster mit buttonclick auf "Akzeptieren und Weiter" automatisch schließen 


            IWebDriver driver = new ChromeDriver("C:\\Users\\holzn\\.cache\\selenium\\chromedriver\\win64");

            //driver.Url = "https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008";
            driver.Url = "https://www.boerse.de";

            Thread.Sleep(10000);
            //  Bei Hauptfenster funktionieren xpath und full xpath, aber man muss natürlich vorher die Unterfenster händisch zumachen
            //driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[1]/div[5]/div/div/div/div/div[2]/span[1]")).Click();
            //driver.FindElement(By.XPath("//*[@id='slideToggleContentAktien']")).Click();
                       
            driver.SwitchTo().Frame("sp_message_iframe_1042969");
            driver.FindElement(By.XPath("//*[@id='notice']/div[3]/div[1]/div[2]/button")).Click();
            //  driver.FindElement(By.XPath("/ html / body / div[2] / div / div / div / div[3] / div[1] / div[2] / button")).Click();

            Thread.Sleep(10000);

 
    }
    /*IWebDriver driver;
        driver = new ChromeDriver("C:\\Users\\holzn\\.cache\\selenium\\chromedriver\\win64");
        //driver.Url = "https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008";
        driver.Url = "https://toolsqa.com/handling-alerts-using-selenium-webdriver/";
        string BaseWindow = driver.CurrentWindowHandle;
        Ausgabe.Text += BaseWindow + "\n";
        Thread.Sleep(5000);
        driver.Manage().Window.Maximize();
        driver.FindElement(By.XPath("//*[@id='content']/p[4]/button")).Click();
        IAlert simpleAlert = driver.SwitchTo().Alert();

        Thread.Sleep(5000);
        driver.Quit();
        //driver.Close();
        //IList<string> totWindowHandles = new List<string>(driver.WindowHandles);
        //Ausgabe.Text= totWindowHandles[0].ToString();
        //Ausgabe.Text= totWindowHandles[1].ToString();
        /*

        ReadOnlyCollection<string> handles = driver.WindowHandles;

        foreach (string handle in handles)
        {
            Ausgabe.Text += handle + "\n";
            Boolean a = driver.SwitchTo().Window(handle).Url.Contains("Main");
            if (a == true)
            {
              //  InitialSetting.driver.SwitchTo().Window(handle);
              //  break;
            }
        }*/

    }
    }

