using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;


namespace Test_import_excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files| *.xls;*.xlsx;*.xlsm";
            if (ope.ShowDialog() == DialogResult.Cancel)
                return;
            FileStream stream = new FileStream(ope.FileName, FileMode.Open);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            DataClasses1DataContext conn = new DataClasses1DataContext();
            foreach (DataTable table in result.Tables)
            {
                foreach (DataRow dr in table.Rows)
                {
                    Test addtabl = new Test()
                    {
                        test_id=Convert.ToString(dr[0]),
                        test_name=Convert.ToString(dr[1]),
                        test_surname=Convert.ToString(dr[2]),
                        test_age=Convert.ToInt32(dr[3]),                        
                    };
                    conn.Tests.InsertOnSubmit(addtabl);
                }
            }
            conn.SubmitChanges();
            excelReader.Close();
            stream.Close();
            MessageBox.Show("lyox lyava");
        }
    }
}
