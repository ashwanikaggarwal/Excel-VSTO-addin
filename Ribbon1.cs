using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelAddIn1.CurrencyService;
using System.Reflection;
using System.Xml;
using System.Windows.Forms;
using System.Data.Sql;
using System.Net;
using System.Data.OleDb;
using Microsoft.Office.Interop.Access;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        //változók
        private string publicStartDate = "";
        private string publicEndDate = "";
        //listák
        private readonly List<string> filesList = new List<string>();
        private readonly List<string> currency = new List<string>();
        private readonly List<decimal> currencyunits = new List<decimal>();
        private readonly List<decimal> currentexchangerates = new List<decimal>();
        private readonly List<decimal> exchangerates = new List<decimal>();
        private readonly List<string> info = new List<string>();
        //meghívom és egy listába teszem a string fájlokat
        public void GetInfo(MNBArfolyamServiceSoapClient client)
        {
            GetInfoRequestBody Gi = new GetInfoRequestBody();
            if (client != null)
            {
                try
                {
                    GetInfoResponseBody getInfoResponseBody = client.GetInfo(Gi); filesList.Add(getInfoResponseBody.GetInfoResult);
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.GetType().FullName + ex.Message); }
            }
        }
        public void GetCurrencies(MNBArfolyamServiceSoapClient client)
        {
            GetCurrenciesRequestBody Curr = new GetCurrenciesRequestBody();
            if (client != null)
            {
                try
                {
                    GetCurrenciesResponseBody getCurrenciesResponseBody = client.GetCurrencies(Curr); filesList.Add(getCurrenciesResponseBody.GetCurrenciesResult);
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.GetType().FullName + ex.Message); }
            }
        }
        public void GetCurrencyUnits(string currencyNames, MNBArfolyamServiceSoapClient client)
        {
            GetCurrencyUnitsRequestBody CurrU = new GetCurrencyUnitsRequestBody
            {
                currencyNames = currencyNames
            };
            if (client != null)
            {

                try
                {
                    GetCurrencyUnitsResponseBody getCurrencyUnitsResponseBody = client.GetCurrencyUnits(CurrU); filesList.Add(getCurrencyUnitsResponseBody.GetCurrencyUnitsResult);
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.GetType().FullName + ex.Message); }
            }
        }
        public void GetCurrentExchangeRates(MNBArfolyamServiceSoapClient client)
        {
            GetCurrentExchangeRatesRequestBody CExchange = new GetCurrentExchangeRatesRequestBody();
            if (client != null)
            {
                try
                {
                    GetCurrentExchangeRatesResponseBody getCurrentExchangeRatesResponseBody = client.GetCurrentExchangeRates(CExchange); filesList.Add(getCurrentExchangeRatesResponseBody.GetCurrentExchangeRatesResult);
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.GetType().FullName + ex.Message); }
            }
        }
        public void GetExchangeRates(string startDate, string endDate, string currencyNames, MNBArfolyamServiceSoapClient client)
        {
            publicStartDate += startDate;
            publicEndDate += endDate;
            GetExchangeRatesRequestBody Exchange = new GetExchangeRatesRequestBody
            {
                startDate = startDate,
                endDate = endDate,
                currencyNames = currencyNames
            };
            if (client != null)
            {
                try
                {
                    GetExchangeRatesResponseBody getExchangeRatesResponseBody = client.GetExchangeRates(Exchange); filesList.Add(getExchangeRatesResponseBody.GetExchangeRatesResult);
                }
                catch (InvalidOperationException ex) { MessageBox.Show(ex.GetType().FullName + ex.Message); }
            }
        }
        //minden aktuális fájlt betöltöm és tovább passzolom
        public void DocumentLoadIn(Excel.Worksheet worksheet)
        {
            XmlDocument document = new XmlDocument
            {
                XmlResolver = null
            };

            for (int j = 1; j - 1 < filesList.Count; j++)
            {
                StringReader stringreader = new StringReader(filesList[j - 1]);
                XmlReader reader = XmlReader.Create(stringreader, new XmlReaderSettings { });
                if (reader == null) { reader.Dispose(); } else {
                    document.Load(reader);
                    LoadMNBIn(document, j, worksheet);
                    reader.Dispose();
                }
                
            }
        }
        //a dokumentum tages elemeit nézem és ha elem, akkor az éppen aktuális betöltött fájl szerint listába teszem az elemek tartalmát 
        //illetve átadom a kiirató metódusnak
        public void LoadMNBIn(XmlDocument document,int j, Excel.Worksheet worksheet)
        {

            if (document != null)
            {
                foreach (XmlNode child in document.DocumentElement.ChildNodes)
                {
                    for (int i = 0; i < child.ChildNodes.Count; i++)
                    {
                        if (child.NodeType == XmlNodeType.Element)
                        {
                            switch (j)
                            {
                                case 1: info.Add(child.ChildNodes[i].InnerText); WriteOut(j, worksheet); break;
                                case 2: currency.Add(child.ChildNodes[i].InnerText);WriteOut(j, worksheet); break;
                                case 3: currencyunits.Add(decimal.Parse(child.ChildNodes[i].InnerText));  WriteOut(j,  worksheet); break;
                                case 4: currentexchangerates.Add(decimal.Parse(child.ChildNodes[i].InnerText));  WriteOut(j, worksheet); break;
                                case 5: exchangerates.Add(decimal.Parse(child.ChildNodes[i].InnerText));  WriteOut(j,  worksheet); break;                                
                                default:break;
                            }
                        }

                    }
                }
            }
        }
        //beállítom, hogy automatikusan elférjen és az aktuális fájlt kiiratom a neki szánt oszlopba és sorokba
        public void WriteOut(int j, Excel.Worksheet worksheet)
        {
            
            if (worksheet != null)
            {
                worksheet.Columns.AutoFit();
                switch (j)
                {
                    case 1: for (int i = 1; i - 1 < info.Count; i++){worksheet.Cells[i, j] = info[i - 1];} break;
                    case 2: for (int i = 1; i - 1 < currency.Count; i++) { worksheet.Cells[i+2, j] = currency[i - 1]; } break;
                    case 3: for (int i = 1; i - 1 < currencyunits.Count; i++) { worksheet.Cells[i+2, j] = currencyunits[i - 1]; } break;
                    case 4: for (int i = 1; i - 1 < currentexchangerates.Count; i++) { worksheet.Cells[i+2, j] = currentexchangerates[i - 1]; } break;
                    case 5:worksheet.Cells[1, j] = publicStartDate; worksheet.Cells[2, j] = publicEndDate; for (int i = 1; i - 1 < exchangerates.Count; i++) { worksheet.Cells[i+2, j] = exchangerates[i - 1]; } break;
                    default: break;
                }
            }
        }
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        //gombnyomásra az aktív könyv első "oldalát" idézem be illetve a szolgáltatás cliensét is és meghívom a folyamatait
        private void MNBbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Workbook workbook1 = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet;

            worksheet = (Excel.Worksheet)workbook1.Worksheets.get_Item(1);

            MNBArfolyamServiceSoapClient client = new MNBArfolyamServiceSoapClient();

            GetInfo(client);
            GetCurrencies(client);
            GetCurrencyUnits("HUF,CHF,EUR,AUD,YUD",client);
            GetCurrentExchangeRates(client);
            GetExchangeRates("2015.01.01", "2020.04.01", "EUR", client);


            client.Close();

            DocumentLoadIn(worksheet);//meghívom a betöltést
        }

    }
}