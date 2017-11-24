using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 创意研发室_待出库统计
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        NPExcel npe = new NPExcel();
        DataTable MasterTable = new DataTable();
        private string GetFileName()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XLS文件(*.xls;*.xlsx)|*.xls;*.xlsx";
            string OpenFile= Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.InitialDirectory = OpenFile;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            else
                return "获取失败";
        }
        private string GetSaveFile()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XLS文件(*.xls)|*.xls";
            string SaveFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            sfd.InitialDirectory = SaveFile;
            if(sfd.ShowDialog()==DialogResult.OK)
            {
                return sfd.FileName;
            }
            return "获取失败";
        }
        private void AddMasterTable(string ExcelFileName)
        {
            npe.LoadFile(ExcelFileName);
            DataTable tempTable = npe.ExcelToDataTable();
            foreach(DataRow dr in tempTable.Rows)
            {
                string CPMC = dr["产品名称"].ToString();
                string CPBH = dr["产品编号"].ToString();
                string CPCount = dr["数量(枚)"].ToString();
                if (CPBH != "")
                {
                    string[] tempCPBH = CPBH.Split('-');
                    string ReplaceCPBH = tempCPBH[3];
                    bool Flag = false;//为False 则主表中没有该数据，Ture则为有该数据不添加了
                    foreach (DataRow MasterDr in MasterTable.Rows)
                    {
                        if (MasterDr["产品编号"].ToString() == ReplaceCPBH)
                        {
                            Flag = true;
                        }
                    }
                    if (Flag == false)
                    {
                        DataRow MasterDr = MasterTable.NewRow();
                        MasterDr["产品名称"] = CPMC;
                        MasterDr["产品编号"] = ReplaceCPBH;
                        MasterDr["数量"] = CPCount;
                        MasterTable.Rows.Add(MasterDr);
                    }
                }
            }
        }
        private void SaveExcel(string FileName,DataTable SaveTable)
        {
            npe.DataTableToExcel(SaveTable, FileName);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string FileName = GetFileName();
            if(FileName != "获取失败")
            {
                AddMasterTable(FileName);
                dataGridView1.DataSource = MasterTable;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MasterTable.Columns.Add("产品编号");
            MasterTable.Columns.Add("产品名称");
            MasterTable.Columns.Add("数量");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string FileName = GetSaveFile();
            if(FileName!="获取失败")
            {
                SaveExcel(FileName, MasterTable);
                MessageBox.Show("保存成功");
            }
        }
    }
}
