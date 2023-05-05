using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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

namespace JsonToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void btnOpen_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "JSON数据文件|*.json";
            var dr = openFileDialog1.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                lbPath.Text = openFileDialog1.FileName;
                rtbJson.Text = File.ReadAllText(lbPath.Text.Trim());
            }
        }
        private void btnJsonToExcel_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(rtbJson.Text))
            {
                var json = rtbJson.Text;
                //var jobj=JsonConvert.DeserializeObject(json);
                JArray items = JArray.Parse(json);
                ToExcel(items, lbPath.Text);
            }
        }
        private void ToExcel(JArray items, string filePath)
        {
            DataSet ds = new DataSet();
            DataTable sheet = new DataTable("sheet");
            string inerIdColName = "__id_";

            int line = 0;
            List<string> propList = new List<string>();
            //propList.Add(inerIdColName);
            //sheet.Columns.Add(inerIdColName, typeof(int));
            sheet.Columns.Add(inerIdColName);

            foreach (JObject item in items)
            {
                if (line++ == 0)
                {
                    int headCol = 0;
                    foreach (var prop in item.Properties())
                    {
                        propList.Add(prop.Name);
                        //sheet.Columns.Add(prop.Name, prop.Value.GetType());
                        sheet.Columns.Add(prop.Name);
                    }
                }

                DataRow dataRow = sheet.NewRow();
                dataRow[inerIdColName] = $"{line}";
                foreach (var propName in propList)
                {
                    dataRow[propName] = item.GetValue(propName);
                }
                sheet.Rows.Add(dataRow);
            }

            ds.Tables.Add(sheet);

            dgv.DataSource = ds;
            dgv.DataMember = "sheet";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(rtbJson.Text))
            {
                var json = rtbJson.Text;
                //var jobj=JsonConvert.DeserializeObject(json);
                JArray items = JArray.Parse(json);
                var xlsPath=SaveExcel(items, lbPath.Text);

                MessageBox.Show($"保存{xlsPath}成功");
            }
        }

        private string SaveExcel(JArray items, string filePath)
        {
            HSSFWorkbook wk = new HSSFWorkbook();
            ICellStyle style11 = wk.CreateCellStyle();
            //创建一个Sheet  
            ISheet sheet = wk.CreateSheet("sheet1");

            int line = 0;
            List<string> propList = new List<string>();
            foreach (JObject item in items)
            {
                if (line == 0)
                {
                    IRow headRow = sheet.CreateRow(0);
                    int headCol = 0;
                    foreach (JProperty prop in item.Properties())
                    {
                        ICell cellHeader = headRow.CreateCell(headCol++);
                        cellHeader.SetCellValue(prop.Name);
                        propList.Add(prop.Name);
                    }
                }

                line++;
                IRow dataRow = sheet.CreateRow(line);
                int dataCol = 0;
                foreach (var prop in propList)
                {
                    ICell dataCell = dataRow.CreateCell(dataCol++);
                    var dataValue = item.GetValue(prop);
                    var dt = dataValue.Type;
                    if (dt == JTokenType.String)
                    {
                        string v = (string)dataValue;
                        dataCell.SetCellValue(v);
                    }
                    else if (dt == JTokenType.Boolean)
                    {
                        bool v = (bool)dataValue;
                        dataCell.SetCellValue(v);
                    }
                    else if (dt == JTokenType.Date)
                    {
                        DateTime v = (DateTime)dataValue;
                        dataCell.SetCellValue(v);
                    }
                    else if (dt == JTokenType.Integer)
                    {
                        int v = (int)dataValue;
                        dataCell.SetCellValue(v);
                    }
                    else if (dt == JTokenType.Float)
                    {
                        float v = (float)dataValue;
                        dataCell.SetCellValue(v);
                    }
                    else
                    {
                        string v = dataValue.ToString();
                        dataCell.SetCellValue(v);
                    }
                }
            }

            for (int i = 0, cnt = propList.Count(); i <= cnt; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            FileInfo fi = new FileInfo(filePath);
            var outPath = Path.Combine(fi.DirectoryName, $"{fi.Name.Replace(".json", "")}.xls");
            using (FileStream fs = File.OpenWrite(outPath))
            {
                wk.Write(fs);//向打开的这个xls文件中写入并保存。  
            }
            return outPath;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
