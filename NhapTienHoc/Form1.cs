using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace NhapTienHoc
{
    public partial class Form1 : Form
    {
        DataTableCollection tables; 
        public Form1()
        {
            InitializeComponent();
        }

        private void CbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tables[cbSheet.SelectedItem.ToString()];
            if(dt != null)
            {
                List<DuLieuDuaLenDataBase> listDuLieuDuaLenDataBase = new List<DuLieuDuaLenDataBase>();
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    DuLieuDuaLenDataBase obj = new DuLieuDuaLenDataBase();
                    obj.maHS = dt.Rows[i]["Mã học sinh"].ToString();
                    obj.maGV = dt.Rows[i]["Mã giáo viên"].ToString();
                    obj.thang = dt.Rows[i]["Tháng"].ToString();
                    obj.soTien = dt.Rows[i]["Tổng số tiền"].ToString();
                    obj.trangThai = dt.Rows[i]["Trạng thái nộp"].ToString();
                    obj.maLop = dt.Rows[i]["Mã lớp"].ToString();
                    listDuLieuDuaLenDataBase.Add(obj);
                }
                dataGridView1.DataSource = listDuLieuDuaLenDataBase;
            }
        }
        
        private void BtnChooseFileExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = ofd.FileName;
                    try
                    {
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                }); ;
                                tables = result.Tables;
                                cbSheet.Items.Clear();
                                foreach (DataTable table in tables)
                                {
                                    cbSheet.Items.Add(table.TableName);
                                }
                            }
                        }
                    }catch(Exception ex)
                    {
                        MessageBox.Show("Xin hãy đóng file Excel", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }
        
        private void BtnImportToDB_Click(object sender, EventArgs e)
        {
            try
            {
                Insert(dataGridView1.DataSource as List<DuLieuDuaLenDataBase>);
                MessageBox.Show("Finished");
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Insert(List<DuLieuDuaLenDataBase> list)
        {
            DapperPlusManager.Entity<DuLieuDuaLenDataBase>().Table("ThongTinNopTien");
            using (IDbConnection db = new SqlConnection("Server=125.212.201.52;Database=NFC;User Id=coneknfc;Password=Conek@123;"))
            {
                db.BulkInsert(list);
            }
        }
    }
}
