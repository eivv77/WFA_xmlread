using Grpc.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WFA_xmlread
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string theDate;

        public void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            theDate = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            listBox1.Items.Clear();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            label1.Text = theDate;
            /*ServicePointManager.SecurityProtocol =
                (SecurityProtocolType)3072 | // TLS 1.2
                (SecurityProtocolType)768 | // TLS 1.1
                (SecurityProtocolType)192;   // TLS 1.0
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Ssl3 |
                SecurityProtocolType.Tls |
                SecurityProtocolType.Tls11 |
                SecurityProtocolType.Tls12 |
                SecurityProtocolType.Tls13;
            ServicePointManager.ServerCertificateValidationCallback = (snder, cert, chain, error) => true;
            ServicePointManager.Expect100Continue = true;
            string url = $"https://www.cbar.az/currencies/{theDate}.xml";
            // use the XmlTextReader to get the xml at the ul
            XmlTextReader reader = new XmlTextReader(url);
            string sp = "";
            int k = 0;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Text)
                {

                    k++;
                    sp += reader.Value + " ";
                    if (k % 3 == 0)
                    {
                        listBox1.Items.Add(sp);

                        sp = "";
                    }
                }
            }*/
            XmlDocument doc = new XmlDocument();
            doc.Load($"https://www.cbar.az/currencies/{theDate}.xml");
            XmlElement root = doc.DocumentElement;
            XmlNodeList nodes = root.SelectNodes("/ValCurs/ValType/Valute/Value");
            var usdNode = nodes[4];
            var euroNode= nodes[5];
            var rublNode = nodes[38];
            listBox1.Items.Add(" USD " + usdNode.InnerText);
            listBox1.Items.Add(" EUR " + euroNode.InnerText);
            listBox1.Items.Add(" RUB " + rublNode.InnerText);

        }

        

        private void button3_Click(object sender, EventArgs e)
        {
            string ExcelFileLocation = ($@"D:\codding\Currency_{theDate}.xlsx");
            Microsoft.Office.Interop.Excel.Application oApp;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Workbook oBook;


            oApp = new Microsoft.Office.Interop.Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oBook.Worksheets.get_Item(1);

            int i = 0;
            i++;

            for (int j = 0; j < listBox1.Items.Count; j++)
            {
                oSheet.Cells[j+1, 1] = listBox1.Items[j].ToString().Split(new char[] { ' ' });
                oSheet.Cells[j+1, 2] = listBox1.Items[j].ToString().Split(new char[] { ' ' });
            }

            foreach (var s in oSheet.Cells)
            {
                Console.WriteLine(s);
            }
            oBook.SaveAs(ExcelFileLocation);
            oBook.Close();
            oApp.Quit();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            XmlReader file;
            file = XmlReader.Create($"https://www.cbar.az/currencies/{theDate}.xml", new XmlReaderSettings());
            DataSet ds = new DataSet();
            ds.ReadXml(file);
            dataGridView1.DataSource = ds.Tables[2];
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns.Add("Column", "Tarix");

        }
    }
}