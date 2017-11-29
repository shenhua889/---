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
        List<int> DataGridViewColumnColour = new List<int>();
        private string GetFileName()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XLS文件(*.xls;*.xlsx)|*.xls;*.xlsx";
            string OpenFile= Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.InitialDirectory = OpenFile; if (ofd.ShowDialog() == DialogResult.OK)
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
        private void UpdateDataGridView()
        {
            dataGridView1.DataSource = MasterTable;
            
            for(int i =0;i<MasterTable.Rows.Count;i++)
            {
                for(int j =0;j<dataGridView1.ColumnCount;j++)
                {
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.White;
                }
                if(MasterTable.Rows[i]["更改"].ToString()=="1")
                {
                    for(int j=0;j<dataGridView1.Columns.Count;j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.MistyRose;
                    }
                }
                MasterTable.Rows[i]["更改"] = 0;
            }
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
                    string ReplaceCPBH = tempCPBH[0]+'-'+tempCPBH[1]+'-'+tempCPBH[2]+'-'+ tempCPBH[3];
                    int Flag = -1;//为False 则主表中没有该数据，Ture则为有该数据不添加了
                    for (int i = 0; i < MasterTable.Rows.Count; i++)
                    {
                        if (MasterTable.Rows[i]["产品编号"].ToString() == ReplaceCPBH)
                        {
                             MasterTable.Rows[i]["数量"] = int.Parse(MasterTable.Rows[i]["数量"].ToString()) + int.Parse(CPCount);
                            MasterTable.Rows[i]["更改"] = 1;
                            Flag = 0;
                        }
                    }
                    if (Flag == -1)
                    {
                        DataRow MasterDr = MasterTable.NewRow();
                        MasterDr["产品名称"] = CPMC;
                        MasterDr["产品编号"] = ReplaceCPBH;
                        MasterDr["数量"] = CPCount;
                        MasterDr["更改"] = 1;
                        MasterTable.Rows.Add(MasterDr);
                    }
                }
            }
        }
        private void UpdateMasterTable(string ExcelFileName)
        {
            npe.LoadFile(ExcelFileName);
            DataTable tempTable = npe.ExcelToDataTable();
            foreach(DataRow tempDr in tempTable.Rows)
            {
                string CPMC = tempDr["产品名称"].ToString();
                string CPBH = tempDr["产品编号"].ToString();
                string CPCount = tempDr["实际销售"].ToString();
                if (CPCount == "")
                    continue;
                for(int i =0;i<MasterTable.Rows.Count;i++)
                {
                    if(MasterTable.Rows[i]["产品编号"].ToString()==CPBH)
                    {
                        MasterTable.Rows[i]["数量"] = int.Parse(MasterTable.Rows[i]["数量"].ToString()) - int.Parse(CPCount);
                        MasterTable.Rows[i]["更改"] = 1;
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
                //dataGridView1.DataSource = MasterTable;
                UpdateDataGridView();
                for(int i=0;i<dataGridView1.Columns.Count;i++)
                {
                    DataGridViewColumn column = dataGridView1.Columns[i];
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MasterTable.Columns.Add("产品编号");
            MasterTable.Columns.Add("产品名称");
            MasterTable.Columns.Add("数量");
            MasterTable.Columns.Add("更改");
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

        private void button2_Click(object sender, EventArgs e)
        {
            string ExcelFileName = GetFileName();
            if(ExcelFileName!="获取失败")
            {
                UpdateMasterTable(ExcelFileName);
                //dataGridView1.DataSource = MasterTable;
                UpdateDataGridView();
            }
        }
    }
}
