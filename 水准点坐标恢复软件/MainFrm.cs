using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.DataSourcesRaster;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.SystemUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 水准点坐标恢复软件
{
    public partial class MainFrm : Form
    {

        private int toolnum = -1;//地图工具条索引
        private string ExcelFile = "";//Excel文件路径
        private struct PointStruct
        {
            public string ptname;//点名
            public double ptlg ;//经纬
            public double ptla;//纬度
            public double dist;//距离
            public int rIndex;//行号
        }
        private Hashtable resHstb = null;

        private PointStruct EditPoint = new PointStruct();//当前编辑的记录
        private bool isEditing = false;//是否处于编辑状态
        private List<PointStruct> changedPtList = new List<PointStruct>();//存储更改过的记录
        public MainFrm()
        {
            InitializeComponent();
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;

            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView.DefaultCellStyle = dataGridViewCellStyle1;

            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col4 = new DataGridViewTextBoxColumn();
           // DataGridViewImageColumn col5 = new DataGridViewImageColumn();
            DataGridViewTextBoxColumn col5 = new DataGridViewTextBoxColumn();

            DataGridViewTextBoxCell celltext = new DataGridViewTextBoxCell();
            DataGridViewImageCell cellimage = new DataGridViewImageCell();

            col1.HeaderText = "点名";
            col1.Name = "ptname";

            col2.HeaderText = "经度";
            col2.Name = "lg";
            col2.Width = 70;

            col3.HeaderText = "纬度";
            col3.Name = "la";
            col3.Width = 70;

            col4.HeaderText = "距离";
            col4.Name = "dist";
            col4.Width = 70;

            col5.HeaderText = "纠正";
            col5.Name = "correct";
            col5.Width = 60;

            dataGridView.Columns.Add(col1);
            dataGridView.Columns.Add(col2);
            dataGridView.Columns.Add(col3);
            dataGridView.Columns.Add(col4);
            dataGridView.Columns.Add(col5);
        }

        /// <summary>
        /// 导入Excel数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImportData_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                 ExcelFile = ofd.FileName;
                 backgroundWorker.RunWorkerAsync();
                
            }
        }

            
       /// <summary>
       /// 读取Excel数据
       /// </summary>
       /// <param name="file">数据路径</param>
       /// <returns>数据存入哈希表</returns>
        private Hashtable ReadExcel()
        {
            Hashtable hm = null;
            ExcelHelper ehp = null;
            Workbook wb = null;
            try
            {
                hm = new Hashtable();
                ehp = new ExcelHelper(false);
                wb = ehp.OpenExcel(ExcelFile);
                Worksheet sheet = ehp.OpenExcelDefaultSheet(ExcelFile);

                int rowcount = int.MaxValue;//行数未知，采用无穷大

                List<PointStruct> ptlist = new List<PointStruct>();

                int obslineNum = 0;//记录测线数
              
                /*
                 * 第五行开始为水准点
                 */
                for (int i = 1; i < rowcount; i++)
                {
                    PointStruct ptst = new PointStruct();

                    /*
                     * 初始化
                     */
                    ptst.ptname = string.Empty;
                    ptst.ptlg = double.NaN;
                    ptst.ptla = double.NaN;
                    ptst.dist = double.NaN;
                    ptst.rIndex = i;
                    double lastDistRec = double.NaN;//记录上一个水准点距离
                    ptst.ptname = ehp.GetCellText(sheet, i, 1);//点名
                    double.TryParse(ehp.GetCellText(sheet, i, 2), out ptst.ptlg);//经度
                    double.TryParse(ehp.GetCellText(sheet, i, 3), out ptst.ptla);//纬度
                    try
                    {
                        ptst.dist = double.Parse(ehp.GetCellText(sheet, i, 4));
                    }//距离
                    catch (Exception e)
                    {
                        ptst.dist = double.NaN;
                    }
                    try
                    {
                        lastDistRec = double.Parse(ehp.GetCellText(sheet, i - 1, 4));
                    }//距离
                    catch (Exception e)
                    {
                        lastDistRec = double.NaN;
                    }
                    if (string.IsNullOrEmpty(ptst.ptname))
                        break;

                    /*
                     * 遍历到最后一个点
                     */

                    if (ehp.GetCellText(sheet, i, 1) == "∑")
                    {
                        obslineNum++;
                        hm.Add(obslineNum, ptlist);
                        break;
                    }
                    /*
                     * 距离差为负数表示新测线的起点
                     */

                    if ((ptst.dist-lastDistRec) < 0 && i != 1)
                    {
                        obslineNum++;
                        hm.Add(obslineNum, ptlist);

                        ptlist = new List<PointStruct>();
                        ptlist.Add(ptst);
                    }
                    else
                    {
                        ptlist.Add(ptst);
                    }
                }
                //ehp.CloseWorkBook(wb, false);

                sheet = null;
                ehp.CloseExcelApplication(false);

            }
            catch (Exception ex)
            {
                //ehp.CloseWorkBook(wb, false);
                ehp.CloseExcelApplication(false);

                MessageBox.Show(ex.Message,"错误");
            }


            return hm;

        }

        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (resHstb != null)
                {
                    IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
                    IActiveView pAcitveView = pGra as IActiveView;
                    pGra.DeleteAllElements();


                    if (dataGridView.Rows.Count > 0)
                        dataGridView.Rows.Clear();

                    string key = comboBox.SelectedItem.ToString();
                    key = key.Split('-')[1];

                    List<PointStruct> ptlist = (List<PointStruct>)resHstb[int.Parse(key)];
                    int c = 0;
                    bool isneededStp = false;//是否需要插入起始点
                    PointStruct startps = new PointStruct();
                    startps.ptname = "起始点";
                    startps.ptlg = 0;
                    startps.ptla = 0;
                    startps.dist = 0;
                    foreach (PointStruct pt in ptlist)
                    {
                        if(c==0)
                            if (ptlist[c].dist != 0)
                            {
                                MessageBox.Show("该测线没有起始点！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                //ExcelHelper ehp = null;
                                try
                                {
                                    //ehp = new ExcelHelper(false);
                                    //Workbook wb = ehp.OpenExcel(ExcelFile);
                                    //Worksheet sheet = ehp.OpenExcelDefaultSheet(ExcelFile);

                                    //startps.rIndex = ptlist[c].rIndex;

                                    //ehp.InsertRows(sheet, ptlist[c].rIndex);
                                    //ehp.FillCellText(sheet, ptlist[c].rIndex, 1, startps.ptname);
                                    //ehp.FillCellText(sheet, ptlist[c].rIndex, 2, startps.ptlg.ToString());
                                    //ehp.FillCellText(sheet, ptlist[c].rIndex, 3, startps.ptla.ToString());
                                    //ehp.FillCellText(sheet, ptlist[c].rIndex, 4, startps.dist.ToString());
                                    //ehp.CloseExcelApplication(true);

                                    //DataGridViewRow row0 = new DataGridViewRow();
                                    //int idx = dataGridView.Rows.Add(row0);
                                    //dataGridView.Rows[idx].Cells[0].Value = startps.ptname;
                                    //dataGridView.Rows[idx].Cells[1].Value = startps.ptlg;
                                    //dataGridView.Rows[idx].Cells[2].Value = startps.ptla;
                                    //dataGridView.Rows[idx].Cells[3].Value = startps.dist;
                                    //DataGridViewButtonCell dgvbc0 = new DataGridViewButtonCell();
                                    //dgvbc0.Value = "+";
                                    //dataGridView.Rows[idx].Cells[4] = dgvbc0;

                                    isneededStp = true;
                                }
                                catch (Exception ex)
                                {
                                    //ehp.CloseExcelApplication(false);
                                }
                              
                            }

                        DataGridViewRow row = new DataGridViewRow();
                        int index = dataGridView.Rows.Add(row);
                        dataGridView.Rows[index].Cells[0].Value = pt.ptname;
                        dataGridView.Rows[index].Cells[1].Value = pt.ptlg;
                        dataGridView.Rows[index].Cells[2].Value = pt.ptla;
                        dataGridView.Rows[index].Cells[3].Value = pt.dist;
                        DataGridViewButtonCell dgvbc = new DataGridViewButtonCell();
                        dgvbc.Value = "+";
                        dataGridView.Rows[index].Cells[4] = dgvbc;

                        DrawPointOrLine(pt.ptlg.ToString(), pt.ptla.ToString(), pGra, false);//图上画点

                        c++;

                    }

                    /*ptlist插入起始点*/
                    //if (isneededStp)
                    //    ptlist.Insert(0, startps);
                    /*更新表格样式*/
                    updateDataGridViewCellStyle();
                    pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 纠正按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0)
                    return;
                string ptname = this.dataGridView.Rows[e.RowIndex].Cells["ptname"].Value.ToString();
                string ptlg = this.dataGridView.Rows[e.RowIndex].Cells["lg"].Value.ToString();
                string ptla = this.dataGridView.Rows[e.RowIndex].Cells["la"].Value.ToString();
                string dist = this.dataGridView.Rows[e.RowIndex].Cells["dist"].Value.ToString();

                if (e.ColumnIndex != -1)
                {
                    if (this.dataGridView.Columns[e.ColumnIndex].Name == "correct")//纠正按钮事件
                    {
                        this.axMapControl.CurrentTool = null;
                        this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerPencil;

                        this.EditPoint.ptname = ptname;
                        double.TryParse(ptlg, out this.EditPoint.ptlg);
                        double.TryParse(ptla, out this.EditPoint.ptla);
                        double.TryParse(dist, out this.EditPoint.dist);
                        this.EditPoint.rIndex = e.RowIndex;

                
                    }
                }
                else
                {

                    double Xmin = this.axMapControl.Extent.XMin;
                    double Xmax = this.axMapControl.Extent.XMax;

                    /*
                      * 经纬度完整
                      * 画点
                      */
                    if ((ptlg != string.Empty && ptlg != "0") && (ptla != string.Empty && ptla != "0"))
                    {
                        IElement pEle = null;

                        ESRI.ArcGIS.Geometry.IPoint pt = new PointClass() { X = double.Parse(ptlg), Y = double.Parse(ptla) };

                        IMarkerElement pMakEle = new MarkerElementClass();
                        pEle = pMakEle as IElement;
                        IMarkerSymbol pMakSym = new SimpleMarkerSymbolClass();
                        pMakSym.Size = 5;


                        pMakSym.Color = ColorToIColor(Color.Turquoise);
                        pMakEle.Symbol = pMakSym;
                        pEle.Geometry = pt;

                        AddSelectedElementByGraphicsSubLayer("selectedEleLayer", pEle);
                        axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                    }

                    /*
                     * 只有纬度的情况
                     * 画纬线
                     */
                    if ((ptlg == string.Empty || ptlg == "0") && ptla != string.Empty)
                    {
                        ISegmentCollection pPath = new PathClass();
                        object o = Type.Missing;

                        ESRI.ArcGIS.Geometry.IPoint FromPt = new PointClass() { X = Xmin, Y = double.Parse(ptla) };
                        ESRI.ArcGIS.Geometry.IPoint ToPt = new PointClass() { X = Xmax, Y = double.Parse(ptla) };

                        ILine2 line = new LineClass();
                        line.FromPoint = FromPt;
                        line.ToPoint = ToPt;

                        pPath.AddSegment(line as ISegment, ref o, ref o);

                        IGeometryCollection pPolyline = new PolylineClass();
                        pPolyline.AddGeometry(pPath as IGeometry, ref o, ref o);

                        ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
                        IColor pColor = new RgbColorClass();

                        lineSymbol.Color = ColorToIColor(Color.Red);//颜色  
                        lineSymbol.Style = esriSimpleLineStyle.esriSLSDash; //样式  
                        lineSymbol.Width = 1;

                        ILineElement pLineElement = new LineElementClass();
                        pLineElement.Symbol = lineSymbol;

                        IElement pElement = pLineElement as IElement;
                        pElement.Geometry = pPolyline as IGeometry;

                        AddSelectedElementByGraphicsSubLayer("selectedEleLayer", pElement);
                        axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 地图鼠标移动事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void axMapControl_OnMouseMove(object sender, IMapControlEvents2_OnMouseMoveEvent e)
        {
            CoordinateLabel.Text = "当前坐标 X = " + e.mapX.ToString() + " Y = " + e.mapY.ToString();
        }

        /// <summary>
        /// 地图鼠标按下事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void axMapControl_OnMouseDown(object sender, IMapControlEvents2_OnMouseDownEvent e)
        {

        }

        /// <summary>
        /// 更新datagridview单元格样式
        /// </summary>
        private void updateDataGridViewCellStyle()
        {
            for (int i = 0; i < this.dataGridView.Rows.Count; i++)
            {

                string lgstr = string.Empty, lastr = string.Empty;
                if(this.dataGridView.Rows[i].Cells["lg"].Value!=null)
                {
                    lgstr = this.dataGridView.Rows[i].Cells["lg"].Value.ToString();
                }
                if (this.dataGridView.Rows[i].Cells["la"].Value != null)
                {
                    lastr = this.dataGridView.Rows[i].Cells["la"].Value.ToString();
                }

                if (lgstr!=string.Empty && lastr!=string.Empty)
                {
                    if (lgstr != "0" && lastr != "0")
                    {
                        dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Green;
                       
                    }
                    else if (lgstr != "0" && lastr == "0")
                    {
                        dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Orange;
                        
                    }
                    else if (lgstr == "0" && lastr != "0")
                    {
                        dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Orange;
                       
                    }
                    else if (lgstr == "0" && lastr == "0")
                    {
                        dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                      
                    }
                }
                else if (lgstr == string.Empty && lastr != string.Empty)
                {
                    dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Orange;
                  
                }
                else if (lgstr != string.Empty && lastr == string.Empty)
                {
                    dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Orange;
                   
                }
                else if (lgstr == string.Empty && lastr == string.Empty)
                {
                    dataGridView.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                    
                }
            }

          

        }

        private void ptsTpMapBtn_Click(object sender, EventArgs e)
        {
            //在绘制前，清除mainkMapControl中的任何图形元素
            IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
            IActiveView pAcitveView = pGra as IActiveView;
            pGra.DeleteAllElements();
          
            for (int i = 0; i < this.dataGridView.Rows.Count; i++)
            {
                string lgstr = string.Empty, lastr = string.Empty;
                if (this.dataGridView.Rows[i].Cells["lg"].Value != null)
                {
                    lgstr = this.dataGridView.Rows[i].Cells["lg"].Value.ToString();
                }
                if (this.dataGridView.Rows[i].Cells["la"].Value != null)
                {
                    lastr = this.dataGridView.Rows[i].Cells["la"].Value.ToString();
                }

                DrawPointOrLine(lgstr, lastr, pGra, false);
            }

            pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        private IColor ColorToIColor(Color color)
        {
            IColor pColor = new RgbColorClass();
            pColor.RGB = color.B * 65536 + color.G * 256 + color.R;
            return pColor;
        }

        /// <summary>
        /// 图上画点或线
        /// </summary>
        /// <param name="lgstr">经度</param>
        /// <param name="lastr">纬度</param>
        /// <param name="pEle">IElement</param>
        /// <param name="isSelected">是否选中</param>
        /// <param name="pGra">IGraphicsContainer</param>
        private void DrawPointOrLine(string lgstr, string lastr, IGraphicsContainer pGra,bool isSelected)
        {
            /*
           * 经纬度范围
           */
            double Xmin = this.axMapControl.Extent.XMin;
            double Xmax = this.axMapControl.Extent.XMax;
            IElement pEle =null;
            /*
                * 经纬度完整
                * 画点
                */
            if (lgstr != string.Empty && lastr != string.Empty)
            {
                ESRI.ArcGIS.Geometry.IPoint pt = new PointClass() { X = double.Parse(lgstr), Y = double.Parse(lastr) };

                IMarkerElement pMakEle = new MarkerElementClass();
                pEle = pMakEle as IElement;
                IMarkerSymbol pMakSym = new SimpleMarkerSymbolClass();
                pMakSym.Size = 5;


                pMakSym.Color = isSelected ? ColorToIColor(Color.Turquoise) : ColorToIColor(Color.Blue);
                pMakEle.Symbol = pMakSym;
                pEle.Geometry = pt;
                pGra.AddElement(pEle, 0);
            }

            /*
             * 只有纬度的情况
             * 画纬线
             */
            if ((lgstr == string.Empty || lgstr == "0") && lastr != string.Empty)
            {
                ISegmentCollection pPath = new PathClass();
                object o = Type.Missing;

                ESRI.ArcGIS.Geometry.IPoint FromPt = new PointClass() { X = Xmin, Y = double.Parse(lastr) };
                ESRI.ArcGIS.Geometry.IPoint ToPt = new PointClass() { X = Xmax, Y = double.Parse(lastr) };

                ILine2 line = new LineClass();
                line.FromPoint = FromPt;
                line.ToPoint = ToPt;

                pPath.AddSegment(line as ISegment, ref o, ref o);

                IGeometryCollection pPolyline = new PolylineClass();
                pPolyline.AddGeometry(pPath as IGeometry, ref o, ref o);

                ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
                IColor pColor = new RgbColorClass();

                lineSymbol.Color = ColorToIColor(Color.Green);//颜色  
                lineSymbol.Style = esriSimpleLineStyle.esriSLSDash; //样式  
                lineSymbol.Width = 1;

                ILineElement pLineElement = new LineElementClass();
                pLineElement.Symbol = lineSymbol;

                IElement pElement = pLineElement as IElement;
                pElement.Geometry = pPolyline as IGeometry;

                pGra.AddElement(pElement, 0);

            }
        }

        private void AddSelectedElementByGraphicsSubLayer(string SubLayerName, IElement element)
        {
            IGraphicsLayer sublayer = FindOrCreateGraphicsSubLayer(SubLayerName); 
            IGraphicsContainer gc = sublayer as IGraphicsContainer;
            gc.DeleteAllElements();
            gc.AddElement(element, 0);
            if (element.Geometry.GeometryType == esriGeometryType.esriGeometryPoint)
            {
                IEnvelope pEnvelope = new EnvelopeClass();
                pEnvelope.XMin = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).X - 0.2;
                pEnvelope.XMax = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).X + 0.2;
                pEnvelope.YMin = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).Y - 0.2;
                pEnvelope.YMax = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).Y + 0.2;
                axMapControl.Extent = pEnvelope;
            }
           // axMapControl.FlashShape(element.Geometry);
            //axMapControl.Map.SelectByShape(element.Geometry, null, false); 
         
         
        }  
        private IGraphicsLayer FindOrCreateGraphicsSubLayer(string SubLayerName)
        {
            IMap Map1 = axMapControl.Map;
            IGraphicsLayer gl = Map1.BasicGraphicsLayer;
            ICompositeGraphicsLayer cgl = gl as ICompositeGraphicsLayer;
            IGraphicsLayer sublayer;
            try
            {
                sublayer = cgl.FindLayer(SubLayerName);
                if (sublayer == null)
                {
                    sublayer = cgl.AddLayer(SubLayerName, null);
                }
            }
            catch (Exception ex)
            {

                sublayer = cgl.AddLayer(SubLayerName, null);

            }
            
            return sublayer;
        }
 
        private void axToolbarControl_OnItemClick(object sender, IToolbarControlEvents_OnItemClickEvent e)
        {
            toolnum = e.index;  
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            resHstb = ReadExcel();
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (resHstb != null)
            {
                ArrayList akeys = new ArrayList(resHstb.Keys);
                akeys.Sort(); //按字母顺序进行排序
                comboBox.Items.Clear();
                foreach (object key in akeys)
                {
                    comboBox.Items.Add("测线-" + key.ToString());
                }
                comboBox.SelectedIndex = 0;
            }
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            if (changedPtList.Count > 0)
            {
                ExcelHelper ehp=null;
                try
                {
                    ehp = new ExcelHelper(false);
                    Workbook wb = ehp.OpenExcel(ExcelFile);
                    Worksheet sheet = ehp.OpenExcelDefaultSheet(ExcelFile);

                    int rowcount = int.MaxValue;

                    for (int i = 1; i < rowcount; i++)
                    {
                        string ptname = ehp.GetCellText(sheet, i, 1);//点名

                        foreach (PointStruct ps in changedPtList)
                        {
                            if (ptname == ps.ptname)
                            {
                                ehp.FillCellText(sheet, i, 2, ps.ptlg.ToString());
                                ehp.FillCellText(sheet, i, 3, ps.ptla.ToString());
                            }
                        }

                        /*
                         * 遍历到最后一个点
                         */
                        if (ehp.GetCellText(sheet, i, 1) == "∑")
                        {
                            break;
                        }

                    }
                    //ehp.SaveWorkBook(wb);
                    ehp.CloseExcelApplication(true);

                    MessageBox.Show("保存成功！","成功",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
             
                }
                catch (Exception ex)
                {
                    MessageBox.Show("保存失败！", "失败", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ehp.CloseExcelApplication(false);
                }
            }
            else
            {
                MessageBox.Show("无数据更改", "退出", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string ptname = this.dataGridView.Rows[e.RowIndex].Cells["ptname"].Value.ToString();
            string ptlg = this.dataGridView.Rows[e.RowIndex].Cells["lg"].Value.ToString();
            string ptla = this.dataGridView.Rows[e.RowIndex].Cells["la"].Value.ToString();
            string dist = this.dataGridView.Rows[e.RowIndex].Cells["dist"].Value.ToString();

            this.EditPoint.ptname = ptname;
            double.TryParse(ptlg, out this.EditPoint.ptlg);
            double.TryParse(ptla, out this.EditPoint.ptla);
            double.TryParse(dist, out this.EditPoint.dist);
            this.EditPoint.rIndex = e.RowIndex;

            #region 更新ptlist里的值

            string key = comboBox.SelectedItem.ToString();
            key = key.Split('-')[1];
            List<PointStruct> ptlist = (List<PointStruct>)resHstb[int.Parse(key)];
            for (int i = 0; i < ptlist.Count; i++)
                if (ptlist[i].ptname == this.EditPoint.ptname)
                    ptlist[i] = this.EditPoint;

            #endregion

            updateChangedPtlist(this.EditPoint);
         
            updateDataGridViewCellStyle();
        }


        /// <summary>
        /// 更新changedPtList子集
        /// </summary>
        /// <param name="editpoint"></param>
        private void updateChangedPtlist(PointStruct editpoint)
        {
            if (changedPtList.Count > 0)
            {
                bool isupdate = false;
                for (int i = 0; i < changedPtList.Count; i++)
                {
                    if (changedPtList[i].ptname == editpoint.ptname)
                    {
                        isupdate = true;
                        changedPtList[i] = editpoint;
                    }
                }

                if (!isupdate)
                    changedPtList.Add(this.EditPoint);
            }
            else
                changedPtList.Add(editpoint);
        }



        /// <summary>
        /// 地理配准
        /// </summary>
        /// <param name="pFromPoint">采集点集</param>
        /// <param name="pTPoint">输入点集</param>
        /// <param name="pRaster">栅格图层</param>
        /// <param name="pSr">参考坐标系</param>
        /// <param name="pSaveFile">输出路径</param>
        /// <param name="pType">格式</param>
        /// <returns></returns>
        public bool GeoReferencing(IPointCollection pFromPoint, IPointCollection pTPoint, IRaster pRaster, ISpatialReference pSr, string pSaveFile, string pType)
        {

            try
            {
                IRasterGeometryProc pRasterGProc = new RasterGeometryProcClass();
                pRasterGProc.Warp(pFromPoint, pTPoint, esriGeoTransTypeEnum.esriGeoTransPolyOrder1, pRaster);
                pRasterGProc.Register(pRaster);
                IRasterProps pRasterPro = pRaster as IRasterProps;
                pRasterPro.SpatialReference = pSr;//定义投影
                if (File.Exists(pSaveFile))
                {
                    File.Delete(pSaveFile);
                }
                pRasterGProc.Rectify(pSaveFile, pType, pRaster);//路径和格式（String）
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }

        private void btnGeorefrencing_Click(object sender, EventArgs e)
        {
            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerCrosshair;//配准状态
        }









    }
}
