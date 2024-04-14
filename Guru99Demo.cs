using Microsoft.Win32;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Markup;
using System.IO;
using Aspose.Cells;
using System.Globalization;

namespace Selenium
{
    internal class Guru99Demo
    {
        private MainWindow MW;
        IWebDriver driver;

        List<string> aktienID = new List<string>();
        List<string> AktienNamen = new List<string>();
        List<string>[] AktWerte;
        List<string>[] AktWerteZeit;


        bool inTextbox = true;
        bool inCSV = true;
        bool inXLSX = true;
        public Guru99Demo(MainWindow mW)
        {
            this.MW = mW;
        }

        [SetUp]
        public void startBrowser()
        {
            //C:\Program Files\Google\Chrome
            //driver = new ChromeDriver("D:\\3rdparty\\chrome");
            // C:\Users\holzn\.cache\selenium\chromedriver\win64\120.0.6099.109
//            driver = new ChromeDriver("C:\\Users\\holzn\\.cache\\selenium\\chromedriver\\win64\\120.0.6099.109");
            driver = new ChromeDriver("C:\\Users\\holzn\\.cache\\selenium\\chromedriver\\win64");

        }

        [Test]
        internal void AndersWerteSchreiben()
        {
            // berechnet Endzeit, bis dann mit allen AktienNamen den aktuellen Wert anhängen in AktWerte[ii],wenn noch leer oder anders als letzter Wert 
            // danach gemäß Flags inTextbox(Ausgabefenster),in CSV (C:\C#\DataA1.csv),XLSX schreiben 
            AktWerte = new List<string>[AktienNamen.Count];
            AktWerteZeit = new List<string>[AktienNamen.Count];
            IWebElement link;
            DateTime AktWerteSchreibenBegin = DateTime.Now;
            DateTime Endzeit = AktWerteSchreibenBegin.AddSeconds(600);
            List<string> ListeSchleifenzeit = new List<string>();
            string Trennzeichen = ";";

            for (int ii = 0; ii < AktienNamen.Count; ii++)
            {
                AktWerte[ii] = new List<string>();
                AktWerteZeit[ii] = new List<string>();
            }

            while (DateTime.Now < Endzeit)           
            {
                var Schleifenzeit = DateTime.Now.ToString();
                ListeSchleifenzeit.Add(Schleifenzeit.Substring(Schleifenzeit.Length - 8));

                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    string hs1 = aktienID[ii];
                    link = driver.FindElement(By.XPath("//*[@id='"
                            + hs1 + "_p_bg" + "']/span"));
                    //"//*[@id='DE000A1EWWW0_p_bg']/span"                    
                   

                    bool AbfrageOk = false;
                    while (!AbfrageOk)
                    {
                        try
                        {
                            if (AktWerte[ii].Count == 0 || AktWerte[ii][AktWerte[ii].Count - 1] != link.Text)
                            {
                                AktWerte[ii].Add(link.Text);
                                string hs = DateTime.Now.ToString();
                                hs = hs.Substring(hs.Length - 8);
                                AktWerteZeit[ii].Add(hs);                                
                            }
                            AbfrageOk = true;
                        }
                        catch (Exception e)
                        {
                            MW.Ausgabe.Text += ii.ToString() + "ErrorAbfrage\n";
                            AbfrageOk = true;
                        }
                    }
                }
            }
            DateTime AktWerteSchreibenEnd = DateTime.Now;            

            if (inTextbox)
            {
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    MW.Ausgabe.Text += AktienNamen[ii] + ",";
                    for (int jj = 0; jj < AktWerte[ii].Count; jj++)
                    {
                        //MW.Ausgabe.Text += AktWerte[ii][0]+ "\n";
                        MW.Ausgabe.Text += AktWerteZeit[ii][jj] + ",";
                        MW.Ausgabe.Text += AktWerte[ii][jj] + ",";
                    }
                    MW.Ausgabe.Text += "\n";
                }
                MW.Ausgabe.Text += "AktWerteSchreibenBegin:" + AktWerteSchreibenBegin.ToLongTimeString() + "\n";
                MW.Ausgabe.Text += "AktWerteSchreibenEnd:" + AktWerteSchreibenEnd.ToLongTimeString() + "\n";
            }

            if (inCSV)
            {
                string strFilePath = @"C:\C#\DataA1.csv";
                StringBuilder sbOutput = new StringBuilder();
                string hs1 = "";

                int MaxZeilen = 1;
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    hs1 += Trennzeichen+AktienNamen[ii] + Trennzeichen;
                    if (AktWerte[ii].Count > MaxZeilen)
                        MaxZeilen = AktWerte[ii].Count;
                }
                hs1.Substring(0, hs1.Length - 1);
                sbOutput.AppendLine(hs1);

                
                for (int ii = 0; ii < MaxZeilen; ii++)
                {
                    //hs1 = ListeSchleifenzeit[ii] + Trennzeichen;
                    hs1 = "";
                    for (int jj = 0; jj < AktienNamen.Count; jj++)
                    {
                        if (AktWerte[jj].Count > ii)
                            hs1 += AktWerteZeit[jj][ii] +Trennzeichen + AktWerte[jj][ii] + Trennzeichen;
                        else
                            hs1 += " " + Trennzeichen;
                    }
                    hs1.Substring(0, hs1.Length - 1);
                    sbOutput.AppendLine(hs1);
                }
                // Create and write the csv file
                File.WriteAllText(strFilePath, sbOutput.ToString());
            }

            if (inXLSX)
            {
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cell cell;
                
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    cell = sheet.Cells[0, 2*ii + 1];
                    cell.PutValue(AktienNamen[ii]);
                }


                for (int jj = 0; jj < AktienNamen.Count; jj++)
                    for (int ii = 0; ii < AktWerte[jj].Count; ii++)
                                    
                    {
                        cell = sheet.Cells[ii + 1, 2 * jj];
                        cell.PutValue(AktWerteZeit[jj][ii]);
                        cell = sheet.Cells[ii + 1, 2*jj+1];
                        cell.PutValue(AktWerte[jj][ii]);
                    }
                
                wb.Save(@"C:\C#\ExcelA1.xlsx", SaveFormat.Xlsx);
            }
        }
        internal void AktWerteSchreiben()
        { // berechnet Endzeit, bis dann mit allen AktienNamen den aktuellen Wert anhängen in AktWerte[ii]
          // danach gemäß Flags inTextbox(Ausgabefenster),in CSV (C:\C#\DataA1.csv),XLSX schreiben 
            AktWerte = new List<string>[AktienNamen.Count];
            IWebElement link;
            DateTime AktWerteSchreibenBegin = DateTime.Now;
            DateTime Endzeit = AktWerteSchreibenBegin.AddSeconds(60);
            List<string> ListeSchleifenzeit = new List<string>();
            string Trennzeichen = ";";

            for (int ii = 0; ii < AktienNamen.Count; ii++)
            {
                AktWerte[ii] = new List<string>();
            }
            
            while (DateTime.Now < Endzeit)          
            {
                var Schleifenzeit = DateTime.Now.ToString();                
                ListeSchleifenzeit.Add(Schleifenzeit.Substring(Schleifenzeit.Length - 8));

                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {                   
                    string hs1 = aktienID[ii];
                    link = driver.FindElement(By.XPath("//*[@id='"
                            + hs1 + "_p_bg" + "']/span"));
                    //"//*[@id='DE000A1EWWW0_p_bg']/span"
                    AktWerte[ii].Add(link.Text);
                }
            }
            DateTime AktWerteSchreibenEnd = DateTime.Now;
            string AktDat = AktWerteSchreibenBegin.ToString().Substring(0,8);

            if (inTextbox)
            {
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    MW.Ausgabe.Text += AktienNamen[ii] + ",";
                    for (int jj = 0; jj < AktWerte[ii].Count; jj++)
                        //MW.Ausgabe.Text += AktWerte[ii][0]+ "\n";
                        MW.Ausgabe.Text += AktWerte[ii][jj]+ ",";
                    MW.Ausgabe.Text += "\n";
                }
                MW.Ausgabe.Text += "AktWerteSchreibenBegin:" + AktWerteSchreibenBegin.ToLongTimeString() + "\n";
                MW.Ausgabe.Text += "AktWerteSchreibenEnd:" + AktWerteSchreibenEnd.ToLongTimeString() + "\n";
            }

            if (inCSV)
            {                
                string strFilePath = @"C:\C#\DataA1.csv";
                StringBuilder sbOutput = new StringBuilder();
                                
                String hs1 = AktDat + Trennzeichen;
                
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    hs1 += AktienNamen[ii] + Trennzeichen;
                }
                hs1.Substring(0, hs1.Length - 1);
                sbOutput.AppendLine(hs1);

                int AnzZeilen = AktWerte[0].Count;
                for (int ii = 0; ii < AnzZeilen; ii++)
                {
                    hs1 = ListeSchleifenzeit[ii] + Trennzeichen;
                    for (int jj=0;jj< AktienNamen.Count;jj++)                 
                        hs1 += AktWerte[jj][ii] + Trennzeichen;
                    hs1.Substring(0, hs1.Length - 1);
                    sbOutput.AppendLine(hs1);
                }
                // Create and write the csv file
                File.WriteAllText(strFilePath, sbOutput.ToString());
            }

            if (inXLSX)
            {                
                Workbook wb = new Workbook();             
                Worksheet sheet = wb.Worksheets[0];
                Cell cell;
                cell = sheet.Cells[0, 0];
                cell.PutValue(AktDat);
                for (int ii = 0; ii < AktienNamen.Count; ii++)
                {
                    cell = sheet.Cells[0, ii+1];
                    cell.PutValue(AktienNamen[ii]);
                }
                

                int AnzZeilen = AktWerte[0].Count;
                for (int ii = 0; ii < AnzZeilen; ii++)
                {
                    cell = sheet.Cells[ii + 1, 0];
                    cell.PutValue(ListeSchleifenzeit[ii]);
                    for (int jj = 0; jj < AktienNamen.Count; jj++)
                    {
                        cell = sheet.Cells[ii + 1, jj + 1];
                        cell.PutValue(AktWerte[jj][ii]);
                    }
                }
                wb.Save(@"C:\C#\ExcelA1.xlsx", SaveFormat.Xlsx);
            }
        }
        public void Hilfs()
        { //Sucht alle "_p_bg" und extrahiert in globale Var. AktienNamen
            driver.Url = "https://www.boerse.de/realtime-kurse/Dax-Aktien/DE0008469008";
            Thread.Sleep(20000);
            String pageSource = driver.PageSource;

            int aktind = 0;
            bool Ende = false;
            string hs1;
            while (!Ende)
            {
                int findInd = pageSource.IndexOf("_p_bg", aktind);
                if (findInd == -1)
                    Ende = true;
                else
                {
                    aktind = findInd + 3;
                    hs1 = pageSource.Substring(findInd - 12, 12);
                    aktienID.Add(hs1);
                    IWebElement link = driver.FindElement(By.XPath("//*[@id='"
                        + hs1 + "_N_bg" + "']/a"));
                    AktienNamen.Add(link.Text);
                }
            }
            /*
            if(inTextbox)
            {
                for (int ii = 0;ii<AktienNamen.Count;ii++)
                {
                    MW.Ausgabe.Text += AktienNamen[ii]+" "+ aktienID[ii]+ "\n";
                }
            }*/
        }
        public void test()
        {
            driver.Url = "http://www.google.co.in";
        }

        [TearDown]
        public void closeBrowser()
        {
            driver.Close();
        }
        internal void test(string url)
        {
            driver.Url = url;
             IWebElement link = driver.FindElement(By.XPath("//*[@id='rt-header']//div[2]/div/ul/li[2]/a"));
            IWebElement link2 = driver.FindElement(By.XPath(".//*[@id='rt-header']//div[2]/div/ul/li[2]/a"));
            MW.Ausgabe.Text = link.Text;
            MW.Ausgabe.Text += link2.Text;
        }
        internal void test1(string url)
        {
            driver.Url = url;

            for (int ii = 0; ii < 10; ii++)
            {
                Thread.Sleep(10000);
                 IWebElement link = driver.FindElement(By.XPath("//*[@id='DE000A1EWWW0_p_bg']/span"));
                
                MW.Ausgabe.Text += link.Text;
            }
        }
        internal void test2(string url)
        {
            driver.Url = url;
            Thread.Sleep(10000);
            String pageSource = driver.PageSource;

            int aktind = 0;
            bool Ende = false;
            string hs1 ;
            while (!Ende)
            {
                int findInd = pageSource.IndexOf("_p_bg", aktind);
                if (findInd == -1)
                    Ende = true;
                else
                {
                    aktind = findInd +3;
                    hs1= pageSource.Substring(findInd-12,17);
                    hs1 = " link = driver.FindElement(By.XPath(\"//*[@id='"
                        +hs1+
                        "']/span\"));";
                    
                    MW.Ausgabe.Text += "\n" + hs1;
                    MW.Ausgabe.Text += "\n" + "MW.Ausgabe.Text += link.Text;";
                }
            }
            MW.Ausgabe.Text += "\nENDE";



            aktind = 0;
            Ende = false;

            while (!Ende)
            {
                int findInd = pageSource.IndexOf("observed=\"true\">", aktind);
                if (findInd == -1)
                    Ende = true;
                else
                {
                    aktind = findInd + 3;
                    int ll = "observed=\"true\">".Length;
                    int EndInd = pageSource.IndexOf("</a>", findInd);

                    hs1 = pageSource.Substring(findInd + ll, EndInd - findInd - ll);

                    MW.Ausgabe.Text += "\n" + hs1;
                }
            }
                
        }

        internal void test3(string url)
        {
            driver.Url = url;
            Thread.Sleep(10000);
            IWebElement link = driver.FindElement(By.XPath("//*[@id='DE000A1EWWW0_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='NL0000235190_N_bg']/a"));
            MW.Ausgabe.Text += link.Text;
            link = driver.FindElement(By.XPath("//*[@id='NL0000235190_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0008404005_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000BASF111_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000BAY0017_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005200000_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005190003_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000A1DAHH0_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
            MW.Ausgabe.Text += "\n";
            
            link = driver.FindElement(By.XPath("//*[@id='DE000CBK1001_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005439004_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0006062144_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000DTR0CK8_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005140008_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005810055_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005552004_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0005557508_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000ENAG999_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
            MW.Ausgabe.Text += "\n";
             link = driver.FindElement(By.XPath("//*[@id='DE0005785604_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0008402215_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0006047004_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0006048432_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0006231004_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007100000_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0006599905_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000A0D9PT0_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0008430026_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000PAG9113_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000PAH0038_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='NL0012169213_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007030009_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
            MW.Ausgabe.Text += "\n";           
             link = driver.FindElement(By.XPath("//*[@id='DE0007037129_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007164600_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007165631_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007236101_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000ENER6Y0_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000SHL1006_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000SYM9999_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE0007664039_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000A1ML7J1_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
             link = driver.FindElement(By.XPath("//*[@id='DE000ZAL1111_p_bg']/span"));
            MW.Ausgabe.Text += link.Text;
            MW.Ausgabe.Text += "AusEnde";

        }

        // link = driver.FindElement(By.XPath("//*[@id='DE000A1EWWW0_p_bg']/span"));
        //DE000A1EWWW0_p_bg
        //NL0000235190

        /*

         */

    }
}
