using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WooCommerceNET;
using WooCommerceNET.WooCommerce.v3;
using WooCommerceNET.WooCommerce.v3.Extension;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WooCommerce1
{
    public partial class Form1 : Form
    {

        public static string store_url; 
        public static string wcKey;
        public static string c_key;
        public static string wc_secret;

        //ck_bb132c28e2228934cac38e86bc620c8ceae5f276 Klucz Klienta/Cunsumer key
        //cs_45614cca635c17d5b3e7c95a1c7a4e3fa6523bd7 Klucz prywatny/Consumer secret //http://127.0.0.1/Aluro1/store/

        /*przykład
         * 
         *  $consumer_key = 'ck_fcedaba8f0fcb0fb4ae4f1211a75da72'; // Add your own Consumer Key here
            $consumer_secret = 'cs_9914968ae9adafd3741c818bf6d704c7'; // Add your own Consumer Secret here
            $store_url = 'http://localhost/&#39;; // Add the home URL to the store you want to connect to here

// Initialize the class
            $wc_api = new WC_API_Client( $consumer_key, $consumer_secret, $store_url );
         * 
         * */

        public static RestAPI rest2;
        public static WCObject wc;
        // WCObject wc2 = new WCObject(rest2);

        //WAPRO DEMO

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            store_url = tbUrlStore.Text;
            wcKey = tbwcKey.Text;
            c_key = tbc_key.Text;
            wc_secret = tbwc_secret.Text;



            rest2 = new RestAPI(store_url, wcKey, wc_secret, requestFilter: RequestFilter);
            wc = new WCObject(rest2);


            bool a = rest2.Debug;

            bool a1 = rest2.IsLegacy;
            string a2 = rest2.oauth_token;
            string url =  rest2.Url;

            richTextBox1.Text += a + " 1 " + a1 + " " + a2 + " " + url ;

            //Task taskA = new Task(connectToWooCommerce);
            //taskA.Start();


        }

        public void connectToWooCommerce()
        {
            store_url = tbUrlStore.Text;
            wcKey = tbwcKey.Text;
            c_key = tbc_key.Text;
            wc_secret = tbwc_secret.Text;

            rest2 = new RestAPI(store_url, wcKey, wc_secret, requestFilter: RequestFilter);
            wc = new WCObject(rest2);
            bool a = rest2.Debug;
            bool a1 = rest2.IsLegacy;
            richTextBox1.Text += a + " 1 " + a1 ;
            //WCObject wc2 = new WCObject(http://127.0.0.1/Aluro1/wp-json/wc/v3/);

            //Please use WooCommerce Restful API Version 3 url for this WCObject. e.g.: http://www.yourstore.co.nz/wp-json/wc/v3/”

            // Create a task and supply a user delegate by using a lambda expression.
            // Start the task.

            // Output a message from the calling thread.


        }
        private void RequestFilter(HttpWebRequest request)
        {
            request.UserAgent = "WooCommerce.NET";
        }

        private void ResponseFilter(HttpWebResponse response)
        {
            var total = int.Parse(response.Headers["X-WP-Total"]);
            var pagecount = int.Parse(response.Headers["X-WP-TotalPages"]);
        }
        
        string connetionString = null;
        SqlConnection connection;
        SqlCommand command;
        string sql = null;
        SqlDataReader dataReader;
        private void button3_Click(object sender, EventArgs e)
        {

            connetionString = "Data Source=" + textBox4.Text + ";Initial Catalog=" + textBox5.Text + ";User ID=" + textBox7.Text + ";Password=" + textBox6.Text + "";
            connection = new SqlConnection(connetionString);
            try
            {
                connection.Open();
                richTextBox3.Text = "Connection Open !";
                connection.Close();
            }
            catch (Exception ex)
            {
                richTextBox3.Text = "Can not open connection";
            }

        }

        private BindingSource bindingSource1 = new BindingSource();
        private SqlDataAdapter dataAdapter = new SqlDataAdapter();

        private void GetData(string selectCommand)
        {
            try
            {
                // Specify a connection string.
                // Replace <SQL Server> with the SQL Server for your Northwind sample database.
                // Replace "Integrated Security=True" with user login information if necessary.
                

                // Create a new data adapter based on the specified query.
                dataAdapter = new SqlDataAdapter(selectCommand, connetionString);

                // Create a command builder to generate SQL update, insert, and
                // delete commands based on selectCommand.
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

                // Populate a new data table and bind it to the BindingSource.
                System.Data.DataTable table = new System.Data.DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };
                dataAdapter.Fill(table);
                bindingSource1.DataSource = table;

                // Resize the DataGridView columns to fit the newly loaded content.
                dataGridView2.AutoResizeColumns(
                    DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
            }
            catch (SqlException)
            {
                MessageBox.Show("To run this example, replace the value of the " +
                    "connectionString variable with a connection string that is " +
                    "valid for your system.");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = bindingSource1;

            sql = richTextBox4.Text;

            GetData(sql);

        }


        Product p2 = new Product()
        {
            name = "product 1",
            description = "product 1",
            price = 123,
            sku = "test",
            sale_price = 123,
            regular_price = 122,
        };

        private void button2_Click(object sender, EventArgs e)
        {
            wc.Product.Add(p2); //docelowo Task
        }

        string SKU = "";
        List<Product> products = new List<Product>();
        Dictionary<string, string> pDic = new Dictionary<string, string>();

        List<ProductCategory> category = new List<ProductCategory>();

        private async void button5_ClickAsync(object sender, EventArgs e)
        {

            products = await wc.Product.GetAll();

            var bindingList = new BindingList<Product>(products);
            var source = new BindingSource(bindingList, null);
            dataGridView1.DataSource = source;
        }


        private void tabPage3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = products;
        }

        private async void button6_ClickAsync(object sender, EventArgs e) //Read category
        {
            listBox1.Items.Clear();
            category = await wc.Category.GetAll();


            dataGridView3.DataSource = category;

            for (int i = 0; i < category.Count; i++)
            {
                listBox1.Items.Add(category[i].id +" "+category[i].name );
            }

            dataGridView3.AutoResizeColumns(
                    DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            dataGridView2.DataSource = bindingSource1;

            sql = "SELECT NAZWA FROM KATEGORIA_ARTYKULU;";

            GetData2(sql);
        }
        private void GetData2(string selectCommand)
        {
            SqlConnection dbConn = new SqlConnection(connetionString);

            string sqlStr = @"SELECT NAZWA FROM KATEGORIA_ARTYKULU;";

            SqlCommand cmd = new SqlCommand(sqlStr, dbConn);
            dbConn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            listBox1.BeginUpdate();
            listBox1.Items.Clear();
            while (rdr.Read())
            {
                listBox2.Items.Add(rdr.GetString(0));
            }
            rdr.Close();
            dbConn.Close();
            listBox1.EndUpdate();
        }


        private void button9_Click(object sender, EventArgs e)
        {
            ProductCategory pc = new ProductCategory();

            string text = listBox2.GetItemText(listBox2.SelectedItem);
            pc.name = text;
            pc.slug = "";
            pc.parent = 0;
            pc.description = "";
            pc.display = "products";

            wc.Category.Add(pc);
        }

        private void button10_Click(object sender, EventArgs e)
        {

            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select file";
            fdlg.InitialDirectory = @"c:\";
            fdlg.FileName = txtFileName.Text;
            fdlg.Filter = "Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = fdlg.FileName;
            }

            Excel excel = new Excel(txtFileName.Text, 1);
            MessageBox.Show(excel.ReadCell(0, 0));

        }
        // Set the grid's column names from row 1.
        private void SetGridColumns(DataGridView dgv,
            object[,] values, int max_col)
        {
            dgvBooks.Columns.Clear();

            // Get the title values.
            for (int col = 1; col <= max_col; col++)
            {
                string title = (string)values[1, col];
                dgv.Columns.Add("col_" + title, title);
            }
        }
        // Set the grid's contents.
        private void SetGridContents(DataGridView dgv,
            object[,] values, int max_row, int max_col)
        {
            // Copy the values into the grid.
            for (int row = 2; row <= max_row; row++)
            {
                object[] row_values = new object[max_col];
                for (int col = 1; col <= max_col; col++)
                    row_values[col - 1] = values[row, col];
                dgv.Rows.Add(row_values);
            }
        }

        private async void button11_Click(object sender, EventArgs e)
        {
            await wc.Category.Delete(15, true, new Dictionary<string, string>());
           // await wc.Category.Delete(15);
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedCells.Count > 0)
            {
                int selectedrowindex = dataGridView3.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = dataGridView3.Rows[selectedrowindex];
                string cellValue = Convert.ToString(selectedRow.Cells["id"].Value);
                richTextBox6.Text = cellValue;
            }
        }

        private void runQuery()
        {
            string queryPhpMyAdmin = richTextBox7.Text;

            if(queryPhpMyAdmin =="")
            {
                MessageBox.Show("Plase insert some sql query");
            }

            string MySQLConnectionString = "datasource=localhost;port=3306;username=aluro1;password=ScisleTajne01!;database=aluro1";
            MySqlConnection databaseConnection = new MySqlConnection(MySQLConnectionString);

            MySqlCommand commandDatabase = new MySqlCommand(queryPhpMyAdmin, databaseConnection);

            try
            {
                databaseConnection.Open();
                MySqlDataReader myReader = commandDatabase.ExecuteReader();
                richTextBox8.Text = "";
                if (myReader.HasRows)
                {
                    while (myReader.Read())
                    {
                        richTextBox8.Text += myReader.GetString(0) + " " + myReader.GetString(1) + " " + myReader.GetString(2) + " " + myReader.GetString(3);
                    }
                }
                else
                {
                    MessageBox.Show("Succes");
                }

            }
            catch(Exception e)
            {
                MessageBox.Show("Query error:" + e.Message);
            }
        }
        private void button12_Click(object sender, EventArgs e)
        {
            runQuery();
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }
    }
}
