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
using ESRI.ArcGIS.Controls;

namespace 水准点坐标恢复软件
{
    public partial class ControlPointFrm : DevExpress.XtraEditors.XtraForm
    {
        public DataTable controlsTb = new DataTable();

        #region 委托事件 
        /// <summary>
        /// 纠正
        /// </summary>
        /// <param name="ct"></param>
        public delegate void RecifyHandler(DataTable ct, string transTypeStr);
        public event RecifyHandler Recify;
        /// <summary>
        /// 删除要素
        /// </summary>
        /// <param name="index"></param>
        public delegate void DeleteElementHandler(int index);
        public event DeleteElementHandler DeleteElement;
        /// <summary>
        /// 全部删除要素
        /// </summary>
        public delegate void DeleteAllElementHandler();
        public event DeleteAllElementHandler DeleteAllElement;

        /// <summary>
        /// 插入控制点
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public delegate void InsertControlPointsHandler(string x,string y,string index);
        public event InsertControlPointsHandler InsertControlPoints;

        /// <summary>
        /// 释放地图操作
        /// </summary>
        public delegate void DisposeMapOperationHandler();
        public event DisposeMapOperationHandler DisposeMapOperation;


        public delegate void SelectedElementHandler(string index);
        public event SelectedElementHandler SelectedElement;

        #endregion

        public ControlPointFrm()
        {
            InitializeComponent();
            controlsTb.Columns.Add("索引", Type.GetType("System.String"));
            controlsTb.Columns.Add("X源", Type.GetType("System.String"));
            controlsTb.Columns.Add("Y源", Type.GetType("System.String"));
            controlsTb.Columns.Add("X地图", Type.GetType("System.String"));
            controlsTb.Columns.Add("Y地图", Type.GetType("System.String"));
            gridControlCps.DataSource = controlsTb;
        }

        public void InsertGridView(string _fromptX,string _fromptY)
        {
            try
            {
                int index = controlsTb.Rows.Count + 1;
                DataRow newRow = controlsTb.NewRow();
                newRow["索引"] = index;
                newRow["X源"] = _fromptX;
                newRow["Y源"] = _fromptY;
                newRow["X地图"] = "0";
                newRow["Y地图"] = "0";
                controlsTb.Rows.Add(newRow);
                gridControlCps.RefreshDataSource();
                InsertControlPoints(_fromptX, _fromptY, index.ToString());
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("插入控制点过程中发生错误：" + ex.Message, "错误");
            }
           
        }

        public void RefreshDataSource()
        {
            gridControlCps.DataSource= controlsTb;
            gridViewCps.RefreshData();
            gridControlCps.RefreshDataSource();
        }
        private void btnDelCps_Click(object sender, EventArgs e)
        {
           int[] handls = gridViewCps.GetSelectedRows();
           for (int i = 0; i < handls.Length;i++ )
           {
               object id = gridViewCps.GetRowCellValue(handls[i], "索引");
               DeleteElement(int.Parse(id.ToString()));
               gridViewCps.DeleteRow(handls[i]);
           }
        }

        private void btnCancle_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ControlPointFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.controlsTb.Rows.Clear();
            this.gridControlCps.RefreshDataSource();

            DeleteAllElement();
            DisposeMapOperation();
           
        }

        private void btnRecify_Click(object sender, EventArgs e)
        {
            if (Recify != null)
            {
                gridViewCps.RefreshData();
                gridControlCps.RefreshDataSource();
                Recify(controlsTb,comboBoxEdit.SelectedItem.ToString());
            }
        }

        private void gridViewCps_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gridViewCps_DoubleClick(object sender, EventArgs e)
        {
           
            try
            {
                if (gridViewCps.FocusedRowHandle == 0)
                {
                    object id = gridViewCps.GetRowCellValue(gridViewCps.FocusedRowHandle, "索引");
                    SelectedElement(id.ToString());
                }
                  
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }  
        }


     
    }
}