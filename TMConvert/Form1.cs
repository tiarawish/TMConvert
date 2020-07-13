using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Net.Http;
using System.Net;
using Ubiety.Dns.Core;
using Newtonsoft.Json.Linq;

namespace TMConvert
{

    public partial class Form1 : Form
    {
        List<string> allCustomersFromImport = new List<string>();
        List<string> allCustomersWithID = new List<string>();
        List<string> allCustomersWithIDID = new List<string>();
        string pathWorkingDir;
        private static HttpClient client = new HttpClient();

        public Form1()
        {
            InitializeComponent();

            listView1.Columns.Add("Kunde", 200);
            listView1.Columns.Add("Artikelnummer / Beschreibung", 130);
            listView1.Columns.Add("Preis", 60);
            listView1.View = View.Details;

            listView2.Columns.Add("Kunde", 200);
            listView2.Columns.Add("Aritkelnummer / Beschreibung", 130);
            listView2.Columns.Add("Existiert", 60);
            listView2.View = View.Details;

            listView3.Columns.Add("Kunde", 100);
            listView3.Columns.Add("Shopware ID", 100);
            listView3.View = View.Details;

            pathWorkingDir = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

            ApiHelper.InitializeClient();
        }

        // UPDATE CSV
        private void button3_Click(object sender, EventArgs e)
        {
            sqlGetUserAddresses();
            sqlGetArticlesDetails();
        }

        // Import from Weclapp
        private async void button4_Click(object sender, EventArgs e)
        {
            disableAllButtons();
            progressBar1.Enabled = true;
            progressBar1.Visible = true;

            List<CustomerModel> allCustomersWithPrices = await LoadCustomer();

            foreach (CustomerModel cm in allCustomersWithPrices)
            {
                if (cm.company != null)
                {
                    allCustomersFromImport.Add(cm.company);

                    foreach (CustomPrice cp in cm.lCustomerPrices)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Text = cm.company;
                        lvi.SubItems.Add(cp.articleNumber);
                        lvi.SubItems.Add(cp.price.ToString());
                        listView1.Items.Add(lvi);                        
                    }
                }
            }

            listView1.View = View.Details;
            label1.Text = listView1.Items.Count.ToString();

            progressBar1.Enabled = false;
            progressBar1.Visible = false;
            enableAllButtons();
        }

        // IMPORT BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //backgroundWorker1.RunWorkerAsync();

            // Import all Excel Data
            openFileDialog1.Title = "Wähle alle zu importierenden Excel Übersichten";
            openFileDialog1.InitialDirectory = pathWorkingDir;
            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                label1.Text = "LOADING";
                disableAllButtons();
                string[] allExcelFiles = openFileDialog1.FileNames;

                //List<ListViewItem> allItems = new List<ListViewItem>();

                foreach (string excelFile in allExcelFiles)
                {
                    string customerName = excelFile.Substring(excelFile.LastIndexOf('\\') + 1, excelFile.Length - excelFile.LastIndexOf('\\') - 6);
                    allCustomersFromImport.Add(customerName);

                    //string[][] csvWeclapp = ReadCSV(csvFile);
                    string[][] excelWeclapp = ReadExcel(excelFile);
                    excelWeclapp[0] = null;

                    foreach (string[] item in excelWeclapp)
                    {
                        if (item != null)
                        {
                            ListViewItem lvi = new ListViewItem();
                            lvi.Text = customerName;
                            lvi.SubItems.Add(item[0]);
                            lvi.SubItems.Add(item[1]);
                            listView1.Items.Add(lvi);
                        }
                    }
                }

                listView1.View = View.Details;
                label1.Text = listView1.Items.Count.ToString();
                enableAllButtons();
            }
            else
            {
                MessageBox.Show("Das war wohl nichts...");
                label1.Text = "Count";
            }
        }

        // MATCH BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            string[] csvCustomerID = ReadCSV1Dim(pathWorkingDir + "\\s_user_addresses.csv", 1);
            string[] csvCustomerName = ReadCSV1Dim(pathWorkingDir + "\\s_user_addresses.csv", 2);

            if (csvCustomerID == null || csvCustomerName == null)
            {
                MessageBox.Show("s_user_addresses.csv nicht gefunden. Bitte erst Update CSV ausführen...");
                return;
            }

            disableAllButtons();
            listView3.Items.Clear();

            try
            {
                foreach (string element in allCustomersFromImport)
                {
                    ListViewItem lvi = new ListViewItem();

                    string res = Array.Find(csvCustomerName, ele => ele.ToLower().Contains(element.ToLower()));
                    if (res != null)
                    {
                        int id = Array.IndexOf(csvCustomerName, res);
                        lvi.Text = element;
                        lvi.SubItems.Add(csvCustomerID[id]);
                        allCustomersWithID.Add(element); // + ";" + id.ToString());
                        allCustomersWithIDID.Add(csvCustomerID[id]);
                    }
                    else
                    {
                        res = Array.Find(csvCustomerName, ele => ele.ToLower().Replace(" ", String.Empty).Contains(element.ToLower()));
                        if (res != null)
                        {
                            int id = Array.IndexOf(csvCustomerName, res);
                            lvi.Text = element;
                            lvi.SubItems.Add(csvCustomerID[id]);
                            allCustomersWithID.Add(element); // + ";" + id.ToString());
                            allCustomersWithIDID.Add(csvCustomerID[id]);
                        }
                        else
                        {
                            lvi.Text = element;
                            lvi.SubItems.Add("NICHT GEFUNDEN");
                        }
                    }

                    listView3.Items.Add(lvi);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            listView3.View = View.Details;
            enableAllButtons();
        }

        // BUTTON EXPORT
        private async void btnExport_Click(object sender, EventArgs e)
        {
            disableAllButtons();
            progressBar1.Enabled = true;
            progressBar1.Visible = true;

            listView2.Items.Clear();

            var result = await Task.Run(() => writeUploadFile());

            label2.Text = result.ToString();

            progressBar1.Enabled = false;
            progressBar1.Visible = false;
            enableAllButtons();
        }

        private Task<int> writeUploadFile()
        {
            string customerID;
            string mid = @""",""0"",""beliebig"",""";
            string end = @""",""0"",""0"",""0""";

            string[][] csvArticlesDetails = ReadCSV(pathWorkingDir + "\\s_articles_details.csv");

            //int counter = Convert.ToInt32(getLastID()) +1;
            int counter = 1;

            string importCSVLinie = null;

            foreach (ListViewItem lvi in listView1.Items)
            {
                //listView3.Items.IndexOfKey(lvi.Text);
                int icustomerID = allCustomersWithID.IndexOf(lvi.Text);

                if (icustomerID >= 0)
                {
                    ListViewItem lvi2 = new ListViewItem();

                    string icustomerIDID = allCustomersWithIDID[icustomerID].ToString();
                    //ListViewItem[] finder = listView3.Items.Find(lvi.Text, false);
                    //customerID = finder[0].SubItems[0].Text;

                    //customerID = listView3.Items[icustomerID].SubItems[1].Text;
                    string current;

                    try
                    {
                        current = lvi.SubItems[1].Text;
                        current = current.Substring(1, current.IndexOf(']') - 1);
                    }
                    catch
                    {
                        current = lvi.SubItems[1].Text;
                    }

                    string articleID = getArticleIDFromPartNumber(current, csvArticlesDetails);

                    lvi2.Text = lvi.Text;
                    lvi2.SubItems.Add(lvi.SubItems[1].Text);

                    if (articleID != null)
                    {

                        importCSVLinie += "\"" + counter.ToString() + "\"," + articleID + ",\"" + icustomerIDID + mid + lvi.SubItems[2].Text.ToString().Replace(',', '.') + end + Environment.NewLine;
                        counter = counter + 1;

                        lvi2.SubItems.Add("JA");
                        lvi2.ForeColor = Color.Green;
                    }
                    else
                    {
                        lvi2.SubItems.Add("NEIN");
                        lvi2.ForeColor = Color.Red;
                    }

                    listView2.Items.Add(lvi2);
                }
            }

            listView2.View = View.Details;
            File.WriteAllText(pathWorkingDir + "\\custom_prices_import.csv", importCSVLinie);
            return Task.FromResult(counter);
        }

        // UPLAOD BUTTON
        private async void button1_Click_1(object sender, EventArgs e)
        {
            disableAllButtons();
            progressBar1.Enabled = true;
            progressBar1.Visible = true;

            var result = await Task.FromResult<string>(sqlPutCustomPrices());

            progressBar1.Enabled = false;
            progressBar1.Visible = false;
            enableAllButtons();
        }

        private void doEverything()
        {
            List<string> ListForImport = new List<string>();

            string[][] csvWeclapp = ReadCSV(@"C:\VSProjects\TMConvert\TMShopWeclapp\ExportSmartSupportOverview.csv");
            string[][] csvArticlesDetails = ReadCSV(@"C:\VSProjects\TMConvert\TMShopWeclapp\s_articles_details.csv");
            string[] articleNumbers = new String[csvWeclapp.GetLength(0)];
            string[] articlePrice = new String[csvWeclapp.GetLength(0)];

            for (int i = 1; i < csvWeclapp.GetLength(0); i++)
            {
                string current = csvWeclapp[i][0];
                articleNumbers[i] = current.Substring(1, current.IndexOf(']') - 1);

                articlePrice[i] = csvWeclapp[i][6];
            }

            string mid = @""",""0"",""beliebig"",""";
            string end = @""",""0"",""0"",""0""";
            string customerID = "6"; // HIER SETZEN

            int counter = 1;
            string importCSVLinie = null;
            for (int i = 1; i < articleNumbers.Length; i++)
            {
                string partnumber = articleNumbers[i];
                if (partnumber != null)
                {
                    string articleID = getArticleIDFromPartNumber(partnumber, csvArticlesDetails);

                    if (articleID != null)
                    {
                        importCSVLinie += "\"" + counter.ToString() + "\"," + articleID + ",\"" + customerID + mid + articlePrice[i].ToString().Replace(',', '.') + end + Environment.NewLine;
                        counter = counter + 1;
                    }
                }
            }

            File.WriteAllText(@"C:\VSProjects\TMConvert\TMShopWeclapp\custom_prices_import.csv", importCSVLinie);
        }

        private string getArticleIDFromPartNumber(string partnumber, string[][] csvArray)
        {
            foreach (string[] line in csvArray)
            {
                if (line[2].Contains(partnumber))
                {
                    return line[1];
                }
            }
            
            return null;
        }

        public string[][] ReadExcel(string file)
        {
            Excel.Application excel = new Excel.Application();
            string[][] result = null;

            if (File.Exists(file))
            {
                Excel.Workbook wkb = excel.Workbooks.Open(file);
                Excel.Worksheet wks = wkb.Worksheets[1];

                Excel.Range startCell = wks.Cells[1, 1];
                Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range xlRange = wks.UsedRange;

                int rows = lastCell.Row;
                int columns = lastCell.Column; 
                
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                result = new string[rowCount][];

                //string[][] results = new string[rowCount][];

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                try
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (i == 1)
                        {
                            string[] lines = new string[2];
                            lines[0] = "Artikelbezeichnung";
                            lines[1] = "Preis";
                            result[0] = lines;
                        }

                        string endDate = xlRange.Cells[i, 4].Value2.ToString();
                        if (endDate.Length < 2) 
                        {
                            string[] lines = new string[2];
                            lines[0] = xlRange.Cells[i, 1].Value2.ToString();
                            lines[1] = xlRange.Cells[i, 7].Value2.ToString();
                            //for (int j = 1; j <= colCount; j++)
                            //{

                            //    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            //    {
                            //        lines[j - 1] = xlRange.Cells[i, j].Value2.ToString();

                            //    }
                            //}
                            result[i - 1] = lines;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(wks);

                wkb.Close();
                Marshal.ReleaseComObject(wkb);
            }

            excel.Quit();
            Marshal.ReleaseComObject(excel);

            return result;
        }

        public string[][] ReadCSV(string file)
        {
            if (File.Exists(file))
            {
                string[] lines = File.ReadAllLines(file);

                string[][] parts = new string[lines.Length][];

                for (int i = 0; i < lines.Length; i++)
                {
                    parts[i] = lines[i].Split(';');
                }

                return parts;
            }

            else
                return null;
        }

        public string[] ReadCSV1Dim(string file, int dimension)
        {
            if (File.Exists(file))
            {
                string[] lines = File.ReadAllLines(file);

                string[] parts = new string[lines.Length];

                for (int i = 0; i < lines.Length; i++)
                {
                    string[] lineSplit = lines[i].Split(';');
                    parts[i] = lineSplit[dimension].Replace("\"",String.Empty);
                }

                return parts;
            }

            else
                return null;
        }

        private string sqlPutCustomPrices()
        {
            string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
            MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
            MySqlCommand command;

            string result = "Error";
            
            try
            {
                mySqlConnection.Open();
                command = mySqlConnection.CreateCommand();
                
                string[] lines = File.ReadAllLines(pathWorkingDir + "\\custom_prices_import.csv");

                deleteAllFromDB();

                foreach (string line in lines)
                {

                    command.CommandText = "INSERT INTO `vio_article_customer_price`(`id`, `articledetailsID`, `customerId`, `from`, `to`, `price`, `pseudoPrice`, `percent`, `deduction_discount`) " +
                    "VALUES (" + line + ")";

                    command.ExecuteNonQuery();
                }

                command.Connection.Close();
                result = "Success";
                MessageBox.Show(result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                mySqlConnection.Close();
            }
            
            return result;
        }

        private void sqlGetArticlesDetails()
        {
            string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
            MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
            MySqlCommand command;
            IDataReader reader;
            string outputFile = "";

            try
            {
                mySqlConnection.Open();
                command = mySqlConnection.CreateCommand();
                command.CommandText = "SELECT * FROM `s_articles`";

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string[] article = new string[3];
                    article[0] = reader[0].ToString(); // ID
                    article[1] = reader[1].ToString(); // Article ID
                    article[2] = reader[2].ToString(); // Product Code

                    outputFile += "\"" + reader[0].ToString() + "\";\"" + reader[1].ToString() + "\";\"" + reader[2].ToString() + "\"" + Environment.NewLine;
                }

                File.WriteAllText(pathWorkingDir + "\\s_articles.csv", outputFile);

                reader.Close();
                command.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                mySqlConnection.Close();
            }
        }

        private void sqlGetUserAddresses()
        {
            string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
            MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
            MySqlCommand command;
            IDataReader reader;
            string outputFile = "";

            try
            {
                mySqlConnection.Open();
                command = mySqlConnection.CreateCommand();
                command.CommandText = "SELECT * FROM `s_user_addresses`";

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string[] article = new string[3];
                    article[0] = reader[0].ToString(); // ID
                    article[1] = reader[1].ToString(); // Article ID
                    article[2] = reader[2].ToString(); // Product Code

                    outputFile += "\"" + reader[0].ToString() + "\";\"" + reader[1].ToString() + "\";\"" + reader[2].ToString() + "\"" + Environment.NewLine;
                }

                File.WriteAllText(pathWorkingDir + "\\s_user_addresses.csv", outputFile);

                reader.Close();
                command.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                mySqlConnection.Close();
            }
        }

        private void disableAllButtons()
        { 
            foreach (Control control in Controls)
            {
                Button b = control as Button;
                if (b != null)
                {
                    b.Enabled = false;
                }
            }
        }

        private void enableAllButtons()
        {
            foreach (Control control in Controls)
            {
                Button b = control as Button;
                if (b != null)
                {
                    b.Enabled = true;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string getLastID()
        {
            string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
            MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
            MySqlCommand command;
            IDataReader reader;
            string result = string.Empty;

            try
            {
                mySqlConnection.Open();
                command = mySqlConnection.CreateCommand();
                command.CommandText = "SELECT COUNT(*) FROM `vio_article_customer_price`";

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    result = reader[0].ToString();
                }

                reader.Close();
                command.Dispose();
                mySqlConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                mySqlConnection.Close();
            }

            return result;
        }

        private void deleteAllFromDB()
        {
            
                string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
                MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
                MySqlCommand command;

                try
                {
                    mySqlConnection.Open();
                    command = mySqlConnection.CreateCommand();
                    command.CommandText = "DELETE FROM `vio_article_customer_price`";
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    
                }
            
        }
        private void deleteFromDB()
        {
            foreach (string id in allCustomersWithIDID)
            {
                string Connectionstring = "server=trademobile.shop;user id=d031fd37;password=aGPBrucu3J2Qh6E9eCgg;port=3306;database=d031fd37"; //SslMode=Preferred
                MySqlConnection mySqlConnection = new MySqlConnection(Connectionstring);
                MySqlCommand command;

                try
                {
                    mySqlConnection.Open();
                    command = mySqlConnection.CreateCommand();
                    command.CommandText = "DELETE FROM `vio_article_customer_price` WHERE `customerId`='" + id + "'";
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                }
            }
        }

        private async Task<List<CustomerModel>> LoadCustomer(int customerid = 0) 
        {
            ApiProcessor apiProcessor = new ApiProcessor();
            resultListCM customerList = await apiProcessor.getCustomers();
            //resultListCP customPricesList = await apiProcessor.getCustomersWithCustomPricesList();
            
            foreach (CustomerModel cm in customerList.result)
            {
                resultListCP prices = await apiProcessor.getCustomPricesForCustomer(cm.id);

                if (prices.result.Count > 0)
                {
                    cm.lCustomerPrices = prices.result;
                }
            }

            customerList.result.RemoveAll(item => item.lCustomerPrices == null);

            return customerList.result;
        }

        //private void apiCallMethod()
        //{
        //    string apiCall = string.Empty;
        //    string customerID = "11662";
        //    List<Customer> lcustomer = new List<Customer>();

        //    //apiCall = "customer?id-eq=" + customerID + "&properties=id,company,customerNumber";
        //    apiCall = "customer?pageSize=500&id-notnull&properties=id,company,customerNumber";
            
        //    string apiCallResult = httpGetWithBase(apiCall);

        //    JObject jsonResults = JObject.Parse(apiCallResult);

        //    foreach (JContainer child in jsonResults.Children())
        //    {
        //        //lcustomer.Add(new Customer { child });
        //    }
        //}

        private string httpGetWithBase(string apiCall)
        {
            RestClient rClient = new RestClient();
            string baseUrl = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/";

            //rClient.endPoint = "https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/articlePrice?properties=articleId,customerId,price&&customerId-eq=11662";
            rClient.endPoint = baseUrl + apiCall;
            
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = "*";
            rClient.userPassword = "7572d2ea-464e-449f-a8e9-39c68ae12fe5";

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            return strResponse;
        }
    }

    public enum httpVerb
    {
        GET,
        POST,
        PUT,
        DELETE
    }

    public enum authenticationType
    {
        Basic,
        NTLM
    }

    public enum autheticationTechnique
    {
        RollYourOwn,
        NetworkCredential
    }

    class RestClient
    {
        public string endPoint { get; set; }
        public httpVerb httpMethod { get; set; }
        public authenticationType authType { get; set; }
        public autheticationTechnique authTech { get; set; }
        public string userName { get; set; }
        public string userPassword { get; set; }


        public RestClient()
        {
            endPoint = string.Empty;
            httpMethod = httpVerb.GET;
        }

        public string makeRequest()
        {
            string strResponseValue = string.Empty;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endPoint);

            request.Method = httpMethod.ToString();

            String authHeaer = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userPassword));
            request.Headers.Add("AuthenticationToken", userPassword);

            HttpWebResponse response = null;

            try
            {
                response = (HttpWebResponse)request.GetResponse();

                //Proecess the resppnse stream... (could be JSON, XML or HTML etc..._

                using (Stream responseStream = response.GetResponseStream())
                {
                    if (responseStream != null)
                    {
                        using (StreamReader reader = new StreamReader(responseStream))
                        {
                            strResponseValue = reader.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                strResponseValue = "{\"errorMessages\":[\"" + ex.Message.ToString() + "\"],\"errors\":{}}";
            }
            finally
            {
                if (response != null)
                {
                    ((IDisposable)response).Dispose();
                }
            }

            return strResponseValue;
        }

    }
}
