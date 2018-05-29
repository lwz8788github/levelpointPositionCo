using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace 水准点坐标恢复软件
{
    public partial class PointDbForm : DevExpress.XtraEditors.XtraForm
    {
        public delegate void ReplaceFromDbHandler(List<PointStruct> ptlist);
        public event ReplaceFromDbHandler ReplaceFromDb;

        public delegate void ReplaceFromExcelHandler(List<PointStruct> ptlist);
        public event ReplaceFromExcelHandler ReplaceFromExcel;

        public delegate void AverageHandler(List<PointStruct> ptlist);
        public event AverageHandler Average;

        public PointDbForm(DataTable datasource)
        {
            InitializeComponent();
            this.gridControlPt.DataSource = datasource;
        }
        public void UpdateDatasource(DataTable datasource)
        {
            try
            {
                this.gridControlPt.DataSource = datasource;
                gridControlPt.RefreshDataSource();
               
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("更新数据库过程中发生错误：" + ex.Message, "错误");
            }

        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                List<PointStruct> ptlist = new List<PointStruct>();
                for (int i = 0; i < this.gridViewPt.RowCount; i++)
                {
                    if (this.gridViewPt.GetDataRow(i)["选择"].ToString() == "True")
                    {
                        PointStruct pts = new PointStruct();
                        pts.ptname = this.gridViewPt.GetRowCellValue(i, "点名").ToString();
                        pts.ptlg = double.Parse(this.gridViewPt.GetRowCellValue(i, "经度").ToString());
                        pts.ptla = double.Parse(this.gridViewPt.GetRowCellValue(i, "纬度").ToString());
                        ptlist.Add(pts);
                    }
                }

                ReplaceFromDb(ptlist);

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExcelToDb_Click(object sender, EventArgs e)
        {
            try
            {
                List<PointStruct> ptlist = new List<PointStruct>();
                for (int i = 0; i < this.gridViewPt.RowCount; i++)
                {
                    if (this.gridViewPt.GetDataRow(i)["选择"].ToString() == "True")
                    {
                        PointStruct pts = new PointStruct();
                        pts.ptname = this.gridViewPt.GetRowCellValue(i, "点名").ToString();
                        pts.ptlg = double.Parse(this.gridViewPt.GetRowCellValue(i, "经度").ToString());
                        pts.ptla = double.Parse(this.gridViewPt.GetRowCellValue(i, "纬度").ToString());
                        ptlist.Add(pts);
                    }
                }

                ReplaceFromExcel(ptlist);

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
            
        }

        private void btnAverage_Click(object sender, EventArgs e)
        {
            try
            {
             List<PointStruct> ptlist = new List<PointStruct>();
                for (int i = 0; i < this.gridViewPt.RowCount; i++)
                {
                    if (this.gridViewPt.GetDataRow(i)["选择"].ToString() == "True")
                    {
                        PointStruct pts = new PointStruct();
                        pts.ptname = this.gridViewPt.GetRowCellValue(i, "点名").ToString();
                        pts.ptlg = double.Parse(this.gridViewPt.GetRowCellValue(i, "经度").ToString());
                        pts.ptla = double.Parse(this.gridViewPt.GetRowCellValue(i, "纬度").ToString());
                        ptlist.Add(pts);
                    }
                }

                Average(ptlist);

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }
    }
}