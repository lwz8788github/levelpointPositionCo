using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Collections;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.SystemUI;
using Microsoft.Office.Interop.Excel;
using ESRI.ArcGIS.DataSourcesRaster;
using ESRI.ArcGIS.Geodatabase;
using System.IO;
using DevExpress.XtraEditors;
using System.Threading;
using stdole;


namespace 水准点坐标恢复软件
{
    /// <summary>
    /// 地图操作类型
    /// </summary>
    public enum MapOperationType
    {
        /// <summary>
        /// 无操作
        /// </summary>
        NoMode = 0,
        /// <summary>
        /// 打开mxd文档
        /// </summary>
        OpenMXD = 1,
        /// <summary>
        /// 添加地图数据
        /// </summary>
        AddMapData = 2,
        /// <summary>
        /// 放大
        /// </summary>
        ZoomIn = 3,
        /// <summary>
        /// 缩小
        /// </summary>
        ZoomOut = 4,
        /// <summary>
        /// 固定比例放大
        /// </summary>
        FixedZoomIn = 5,
        /// <summary>
        /// 固定比例缩小
        /// </summary>
        FixedZoomOut = 6,
        /// <summary>
        /// 漫游
        /// </summary>
        Pan = 7,
        /// <summary>
        /// 前一视图
        /// </summary>
        PreView = 8,
        /// <summary>
        /// 下一视图
        /// </summary>
        NextView = 9,
        /// <summary>
        /// 全图
        /// </summary>
        FullExtent = 10,
        /// <summary>
        /// 测量
        /// </summary>
        Ruler = 11,
        /// <summary>
        ///配准
        /// </summary>
        Rectify = 12,
        /// <summary>
        /// 坐标纠正
        /// </summary>
        Correct =13
    }
    public struct PointStruct
    {
        public string ptname;//点名
        public double ptlg;//经纬
        public double ptla;//纬度
        public string dist;//距离
        public int rIndex;//行号
    }
    public partial class RibbonForm : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        #region 公共变量
        private string ExcelFile = "";//Excel文件路径
       
        private Hashtable resHstb = null;
        private PointStruct EditPoint = new PointStruct();//当前编辑的记录
        private List<PointStruct> changedPtList = new List<PointStruct>();//存储更改过的记录
        private ControlPointFrm cpf = null;
        private PointDbForm pdf = null;
        private MapOperationType currentMapOperationType;
        #endregion

        #region 构造
        public RibbonForm()
        {
            InitializeComponent();
            axToolbarControl.Visible = false;

            //DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            //dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dataGridView.DefaultCellStyle = dataGridViewCellStyle1;

            //DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
            //DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn();
            //DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn();
            //DataGridViewTextBoxColumn col4 = new DataGridViewTextBoxColumn();
            //// DataGridViewImageColumn col5 = new DataGridViewImageColumn();
            //DataGridViewTextBoxColumn col5 = new DataGridViewTextBoxColumn();

            //DataGridViewTextBoxCell celltext = new DataGridViewTextBoxCell();
            //DataGridViewImageCell cellimage = new DataGridViewImageCell();

            ////DataGridViewCellStyle style = new DataGridViewCellStyle();
            ////style.Font = new System.Drawing.Font("宋体", 11);  
          

            //col1.HeaderText = "点名";
            //col1.Name = "ptname";
            //col1.HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9);  

            //col2.HeaderText = "经度";
            //col2.Name = "lg";
            //col2.Width = 70;
            //col2.HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9);  

            //col3.HeaderText = "纬度";
            //col3.Name = "la";
            //col3.Width = 70;
            //col3.HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9);  

            //col4.HeaderText = "距离";
            //col4.Name = "dist";
            //col4.Width = 70;
            //col4.HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9);  

            //col5.HeaderText = "纠正";
            //col5.Name = "correct";
            //col5.Width = 60;
            //col5.HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9);  

            //dataGridView.Columns.Add(col1);
            //dataGridView.Columns.Add(col2);
            //dataGridView.Columns.Add(col3);
            //dataGridView.Columns.Add(col4);
            //dataGridView.Columns.Add(col5);

            
        }
        #endregion

        #region 方法

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

                int rowcount = ehp.GetRowsCount(sheet); // int.MaxValue;//行数未知，采用无穷大

                ShowDialogForm sdf = new ShowDialogForm("提示", "请稍候...", "正在读取数据！", rowcount);

                List<PointStruct> ptlist = new List<PointStruct>();

                int obslineNum = 0;//记录测线数

                /*
                 * 第五行开始为水准点
                 */
                for (int i = 1; i <= rowcount; i++)
                {
                    PointStruct ptst = new PointStruct();

                    /*
                     * 初始化
                     */
                    ptst.ptname = string.Empty;
                    ptst.ptlg = double.NaN;
                    ptst.ptla = double.NaN;
                    ptst.dist = string.Empty;
                    ptst.rIndex = i;
                    //double lastDistRec = double.NaN;//记录上一个水准点距离
                    ptst.ptname = ehp.GetCellText(sheet, i, 1);//点名
                    double.TryParse(ehp.GetCellText(sheet, i, 2), out ptst.ptlg);//经度
                    double.TryParse(ehp.GetCellText(sheet, i, 3), out ptst.ptla);//纬度
                    ptst.dist = ehp.GetCellText(sheet, i, 4);
                   
                    //try
                    //{
                    //    lastDistRec = double.Parse(ehp.GetCellText(sheet, i - 1, 4));
                    //}//距离
                    //catch (Exception e)
                    //{
                    //    lastDistRec = double.NaN;
                    //}
                    if (string.IsNullOrEmpty(ptst.ptname))
                        break;

                    /*
                     * 距离差为负数表示新测线的起点
                     */
                    //if ((ptst.dist - lastDistRec) < 0 && i != 1)
                    //{
                    //    obslineNum++;
                    //    hm.Add(obslineNum, ptlist);

                    //    ptlist = new List<PointStruct>();
                    //    ptlist.Add(ptst);
                    //}
                    //else
                    //{
                        ptlist.Add(ptst);
                    //}

                    /*
                   * 遍历到最后一个点
                   */
                    //if (ehp.GetCellText(sheet, i, 1) == "∑")
                    if (i == rowcount)
                    {
                        obslineNum++;
                        hm.Add(obslineNum, ptlist);
                        break;
                    }

                    sdf.SetCaption("执行进度（" + i.ToString() + "/" + rowcount.ToString() + "）");
                }
                //ehp.CloseWorkBook(wb, false);

                sheet = null;
                ehp.CloseExcelApplication(false);
                sdf.Close();

            }
            catch (Exception ex)
            {
                //ehp.CloseWorkBook(wb, false);
                ehp.CloseExcelApplication(false);

                XtraMessageBox.Show("Excel格式不对，请采用正确格式！", "错误");
            }
            return hm;
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
        private void DrawPointOrLine(string lgstr, string lastr, IGraphicsContainer pGra, bool isSelected)
        {
            /*
           * 经纬度范围
           */
            double Xmin = this.axMapControl.Extent.XMin;
            double Xmax = this.axMapControl.Extent.XMax;
            IElement pEle = null;
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
            //gc.DeleteAllElements();
            gc.AddElement(element, 0);
            
            //if (element.Geometry.GeometryType == esriGeometryType.esriGeometryPoint)
            //{
            //    IEnvelope pEnvelope = new EnvelopeClass();
            //    pEnvelope.XMin = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).X - 0.1;
            //    pEnvelope.XMax = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).X + 0.1;
            //    pEnvelope.YMin = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).Y - 0.1;
            //    pEnvelope.YMax = ((ESRI.ArcGIS.Geometry.IPoint)element.Geometry).Y + 0.1;
            //    axMapControl.Extent = pEnvelope;
            //}
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
                    sublayer =cgl.AddLayer(SubLayerName, null);
                }
            }
            catch (Exception ex)
            {

                sublayer = cgl.AddLayer(SubLayerName, null);

            }

            return sublayer;
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
        public bool GeoReferencing(IPointCollection pFromPoint, IPointCollection pTPoint, IRaster pRaster, ISpatialReference pSr,esriGeoTransTypeEnum egtt)
        {

            try
            {
                IRasterGeometryProc pRasterGProc = new RasterGeometryProcClass();
                pRasterGProc.Warp(pFromPoint, pTPoint, egtt, pRaster);
                pRasterGProc.Register(pRaster);
                IRasterProps pRasterPro = pRaster as IRasterProps;
         
                pRasterPro.SpatialReference = pSr;//定义投影
                //if (File.Exists(pSaveFile))
                //{
                //    File.Delete(pSaveFile);
                //}
                //pRasterGProc.Rectify(pSaveFile, pType, pRaster);//路径和格式（String）
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return true;
        }


        /// <summary>
        /// 执行地图操作
        /// </summary>
        /// <param name="OpType">操作模式</param>
        private void ExcuteMapOperation(MapOperationType OpType)
        {
            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerDefault;
            switch (OpType)
            {
                case MapOperationType.OpenMXD:
                    {
                        axToolbarControl.GetItem(0).Command.OnClick();
                    }
                    break;
                case MapOperationType.AddMapData:
                    {
                        axToolbarControl.GetItem(1).Command.OnClick();
                    }
                    break;
                case MapOperationType.ZoomIn:
                    {

                        ICommand command = new ControlsMapZoomInToolClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            axMapControl.CurrentTool = (ITool)command;
                        }
                    }
                    break;
                case MapOperationType.ZoomOut:
                    {
                        ICommand command = new ControlsMapZoomOutToolClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            axMapControl.CurrentTool = (ITool)command;
                        }
                    }
                    break;

                case MapOperationType.FixedZoomIn:
                    {
                        ICommand command = new ControlsMapZoomInFixedCommandClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            command.OnClick();
                        }
                    }
                    break;
                case MapOperationType.FixedZoomOut:
                    {
                        ICommand command = new ControlsMapZoomOutFixedCommandClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            command.OnClick();
                        }
                    }
                    break;

                case MapOperationType.Pan:
                    {
                        ICommand command = new ControlsMapPanToolClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            axMapControl.CurrentTool = (ITool)command;
                        }
                    }
                    break;
                case MapOperationType.PreView:
                    {

                        ICommand command = new ControlsMapZoomToLastExtentBackCommandClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            command.OnClick();
                        }
                    }
                    break;
                case MapOperationType.NextView:
                    {

                        ICommand command = new ControlsMapZoomToLastExtentForwardCommandClass();
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            command.OnClick();
                        }
                    }
                    break;
                case MapOperationType.FullExtent:
                    {
                        ICommand command = new ControlsMapFullExtentCommandClass();  //根据该下标志获取点击命令
                        command.OnCreate(axMapControl.Object);
                        if (command.Enabled == true)
                        {
                            command.OnClick();
                        }
                    }
                    break;
                case MapOperationType.Ruler:
                    {
                        axToolbarControl.GetItem(2).Command.OnClick();
                    }
                    break;
            }
        }

        /// <summary>
        /// 重画测线
        /// </summary>
        /// <param name="PromptStartPoint">是否提示起始点</param>
        private void DrawLine(bool PromptStartPoint)
        {
            try
            {
                if (resHstb != null)
                {
                    IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
                    IActiveView pAcitveView = pGra as IActiveView;
                    pGra.DeleteAllElements();


                    //for (int i = 0; i < dataGridView.RowCount; i++)
                    //    dataGridView.DeleteRow(i);
               

                    List<PointStruct> ptlist = (List<PointStruct>)resHstb[1];
                    int c = 0;
                    bool isneededStp = false;//是否需要插入起始点
                    PointStruct startps = new PointStruct();
                    startps.ptname = "起始点";
                    startps.ptlg = 0;
                    startps.ptla = 0;
                    startps.dist = "";
                    foreach (PointStruct pt in ptlist)
                    {
                        if (c == 0)
                          
                                //if (PromptStartPoint)
                                //    XtraMessageBox.Show("该测线没有起始点！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                //ExcelHelper ehp = null;
                            try
                            {
                                isneededStp = true;
                            }
                            catch (Exception ex)
                            {
                                //ehp.CloseExcelApplication(false);
                            }

                        //DataGridViewRow row = new DataGridViewRow();
                        //int index = dataGridView.row
                        //dataGridView.Rows[index].Cells[0].Value = pt.ptname;
                        //dataGridView.Rows[index].Cells[1].Value = pt.ptlg;
                        //dataGridView.Rows[index].Cells[2].Value = pt.ptla;
                        //dataGridView.Rows[index].Cells[3].Value = pt.dist;
                        //DataGridViewButtonCell dgvbc = new DataGridViewButtonCell();
                        //dgvbc.Value = "+";
                        //dataGridView.Rows[index].Cells[4] = dgvbc;

                        DrawPointOrLine(pt.ptlg.ToString(), pt.ptla.ToString(), pGra, false);//图上画点

                        c++;

                    }

                   
                    pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion

        #region 事件

     

        /// <summary>
        /// 地图鼠标移动事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void axMapControl_OnMouseMove(object sender, IMapControlEvents2_OnMouseMoveEvent e)
        {
          
            CoordinateLabel.EditValue = "X = " + e.mapX.ToString() + " Y = " + e.mapY.ToString();
        }

        /// <summary>
        /// 地图鼠标按下事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void axMapControl_OnMouseDown(object sender, IMapControlEvents2_OnMouseDownEvent e)
        {
            try
            {
           
                switch (currentMapOperationType)
                {
                    /*编辑状态*/
                    case MapOperationType.Correct:
                        {
                            bool isChanged = false;//是否有更改
                            if (this.EditPoint.ptlg == double.NaN || this.EditPoint.ptlg == 0)
                            {
                                this.EditPoint.ptlg = double.Parse(e.mapX.ToString("f3"));
                                this.dataGridView.SetRowCellValue(this.EditPoint.rIndex, "lg", this.EditPoint.ptlg);
                                isChanged = true;
                            }
                            else
                            {
                                if (DialogResult.OK == XtraMessageBox.Show("是否替换原经度？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
                                {
                                    this.EditPoint.ptlg = double.Parse(e.mapX.ToString("f3"));
                                    this.dataGridView.SetRowCellValue(this.EditPoint.rIndex, "lg", this.EditPoint.ptlg);
                                    isChanged = true;
                                }
                            }
                            if (this.EditPoint.ptla == double.NaN || this.EditPoint.ptla == 0)
                            {
                                this.EditPoint.ptla = double.Parse(e.mapY.ToString("f3"));
                                this.dataGridView.SetRowCellValue(this.EditPoint.rIndex, "la", this.EditPoint.ptla);
                                isChanged = true;
                            }
                            else
                            {
                                if (DialogResult.OK == XtraMessageBox.Show("是否替换原纬度？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
                                {
                                    this.EditPoint.ptla = double.Parse(e.mapY.ToString("f3"));
                                    this.dataGridView.SetRowCellValue(this.EditPoint.rIndex, "la", this.EditPoint.ptla);
                                    isChanged = true;
                                }
                            }

                            if (isChanged)
                            {
                               // updateChangedPtlist(this.EditPoint);
                                UpdatePtlist();
                                DrawLine(false);
                            }

                            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerDefault;//停止编辑状态
                            currentMapOperationType = MapOperationType.NoMode;
                        }
                        break;
                    /*配准状态*/
                    case MapOperationType.Rectify:
                        {
                            cpf.InsertGridView(e.mapX.ToString(), e.mapY.ToString());
                        }
                        break;
                }
            }
            catch (Exception ep)
            {
                XtraMessageBox.Show(ep.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ptsTpMapBtn_Click(object sender, EventArgs e)
        {
            //在绘制前，清除mainkMapControl中的任何图形元素
            IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
            IActiveView pAcitveView = pGra as IActiveView;
            pGra.DeleteAllElements();

            for (int i = 0; i < this.dataGridView.RowCount; i++)
            {
                string lgstr = string.Empty, lastr = string.Empty;
                if (this.dataGridView.GetRowCellValue(i,"lg") != null)
                {
                    lgstr = this.dataGridView.GetRowCellValue(i, "lg").ToString();
                }
                if (this.dataGridView.GetRowCellValue(i, "lg") != null)
                {
                    lastr = this.dataGridView.GetRowCellValue(i,"la").ToString();
                }

                DrawPointOrLine(lgstr, lastr, pGra, false);
            }
            pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            resHstb = null;
            EditPoint = new PointStruct();//当前编辑的记录
            changedPtList = new List<PointStruct>();//存储更改过的记录

            resHstb = ReadExcel();
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            ArrayList akeys = new ArrayList(resHstb.Keys);
            akeys.Sort(); //按字母顺序进行排序

            System.Data.DataTable dt = new System.Data.DataTable("exceltb");
            DataColumn dc1 = new DataColumn("ptname", Type.GetType("System.String")); 
            DataColumn dc2 = new DataColumn("lg", Type.GetType("System.Double")); 
            DataColumn dc3 = new DataColumn("la", Type.GetType("System.Double")); 
            DataColumn dc4 = new DataColumn("dist", Type.GetType("System.String")); 
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            dt.Columns.Add(dc4);
            List<PointStruct> pointtlist = (List<PointStruct>)resHstb[1];
            for (int j = 0; j < pointtlist.Count; j++)
            {
                DataRow dr = dt.NewRow();
                dr["ptname"] = pointtlist[j].ptname;
                dr["lg"] = pointtlist[j].ptlg;
                dr["la"] = pointtlist[j].ptla;
                dr["dist"] = pointtlist[j].dist;
                dt.Rows.Add(dr);
            }

            this.gridControl1.DataSource = dt;
            DrawLine(true);
        }

   
        
        private void btnImportExcel_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Microsoft Excel files(*.xls)|*.xls;*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExcelFile = ofd.FileName;

                this.gridControl1.DataSource = null;
                this.gridControl1.Update();

            
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void btnSaveAndExit_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (this.dataGridView.RowCount == 0)
            {
                XtraMessageBox.Show("尚未导入数据！", "提示");
                    return;
            }
            if (!AccessHelper.ConnectionTest())
            {
                if (XtraMessageBox.Show("数据库连接不成功,将影响到数据对比功能，是否忽略？", "提示") == System.Windows.Forms.DialogResult.Cancel)
                    return;
            }

            ExcelHelper ehp = null;
            try
            {
                ehp = new ExcelHelper(false);
                Workbook wb = ehp.OpenExcel(ExcelFile);
                Worksheet sheet = ehp.OpenExcelDefaultSheet(ExcelFile);

                PointDao pd = new PointDao();

                int rowcount = ehp.GetRowsCount(sheet);

                ShowDialogForm sdf = new ShowDialogForm("提示", "请稍候...", "正在保存数据！", rowcount);
              
                for (int i = 1; i <= rowcount; i++)
                {
                    string ptname = ehp.GetCellText(sheet, i, 1);//点名
                    double lg = 0;
                    double.TryParse(ehp.GetCellText(sheet, i, 2), out lg);//经度
                    double la = 0;
                    double.TryParse(ehp.GetCellText(sheet, i, 3), out la);//经度

                    foreach (PointStruct ps in changedPtList)
                    {
                        if (ptname == ps.ptname)
                        {
                            ehp.FillCellText(sheet, i, 2, ps.ptlg.ToString());
                            ehp.FillCellText(sheet, i, 3, ps.ptla.ToString());
                            ehp.FillCellText(sheet, i, 4, ps.dist.ToString());

                            lg = ps.ptlg;
                            la = ps.ptla;
                        }
                    }
                    if (double.IsNaN(lg) || double.IsNaN(la))
                        continue;
                    if (lg == 0 || la == 0)
                        continue;

                    try
                    {
                        if (ptname.Contains("'"))
                            ptname = ptname.Replace("'", "''");
                        if (!pd.IsExist(ptname))
                        {
                            pd.InserToDb(ptname, lg, la);
                        }
                    }
                    catch (Exception)
                    {
                        
                        //throw;
                    }
                    

                    sdf.SetCaption("执行进度（" + i.ToString() + "/" + rowcount.ToString() + "）");
                }
                sdf.Close();

              
                XtraMessageBox.Show("保存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("保存失败！", "失败", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
               
                ehp.CloseExcelApplication(true);
                //this.Close();
            }

        }

        #region 地图操作事件
       
        private void btnOpenMXD_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.OpenMXD;
            ExcuteMapOperation(MapOperationType.OpenMXD);
        }

        private void btnAddMap_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.AddMapData;
            ExcuteMapOperation(MapOperationType.AddMapData);
        }

        private void btnZoomin_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.ZoomIn;
            ExcuteMapOperation(MapOperationType.ZoomIn);
        }

        private void btnZoomout_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.ZoomOut;
            ExcuteMapOperation(MapOperationType.ZoomOut);
        }

        private void btnFixedZoomin_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.FixedZoomIn;
            ExcuteMapOperation(MapOperationType.FixedZoomIn);
        }

        private void btnFixedZoomout_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.FixedZoomOut;
            ExcuteMapOperation(MapOperationType.FixedZoomOut);
        }

        private void btnFullmap_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.FullExtent;
            ExcuteMapOperation(MapOperationType.FullExtent);
        }

        private void btnRuler_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.Ruler;
            ExcuteMapOperation(MapOperationType.Ruler);
        }

        private void btnPreView_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.PreView;
            ExcuteMapOperation(MapOperationType.PreView);
        }

        private void btnNextView_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.NextView;
            ExcuteMapOperation(MapOperationType.NextView);
        }

        private void btnPan_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentMapOperationType = MapOperationType.Pan;
            ExcuteMapOperation(MapOperationType.Pan);
        }

        #endregion

        private void btnGeopreferencing_ItemClick(object sender, ItemClickEventArgs e)
        {
            this.axMapControl.CurrentTool = null;
            currentMapOperationType = MapOperationType.Rectify;
            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerCrosshair;
            if (cpf == null || cpf.IsDisposed)
            {
                cpf = new ControlPointFrm();
                cpf.Recify += cpf_Recify;
                cpf.DeleteElement += cpf_DeleteElement;
                cpf.DeleteAllElement += cpf_DeleteAllElement;
                cpf.InsertControlPoints += cpf_InsertControlPoints;
                cpf.DisposeMapOperation += cpf_DisposeMapOperation;
                cpf.SelectedElement += cpf_SelectedElement;
                cpf.Show();
            }
            else
                cpf.Activate();
        }

        private void btnCheckDb_ItemClick(object sender, ItemClickEventArgs e)
        {
            pdf = new PointDbForm(GetDtInExcel());
            pdf.ReplaceFromDb += pdf_ReplaceFromDb;
            pdf.ReplaceFromExcel += pdf_ReplaceFromExcel;
            pdf.Average += pdf_Average;
            pdf.Show();
        }

       

        private System.Data.DataTable GetDtInExcel()
        {
            System.Data.DataTable dt = null;
            List<string> ptnamelist = new List<string>();

            try
            {
                for (int i = 0; i < this.dataGridView.RowCount; i++)
                {
                    object obj = this.dataGridView.GetRowCellValue(i, "ptname");
                    if (obj != null)
                    {
                        string ptname = obj.ToString();
                        if (ptname.Contains("'"))
                            ptname = ptname.Replace("'", "''");
                        ptnamelist.Add(ptname);
                    }
                }
                if (ptnamelist.Count == 0)
                    return dt;

                PointDao pd = new PointDao();
                dt = pd.GetPtlistDb(ptnamelist);


                if (dt.Rows.Count > 0)
                {
                    dt.Columns.Add("选择", System.Type.GetType("System.Boolean"));

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string dtptname = dt.Rows[i]["点名"].ToString();
                        double dtlg = 0;
                        double.TryParse(dt.Rows[i]["经度"].ToString(), out dtlg);
                        double dtla = 0;
                        double.TryParse(dt.Rows[i]["纬度"].ToString(), out dtla);

                        for (int j = 0; j < this.dataGridView.RowCount; j++)
                        {
                            try
                            {
                                string ptname = this.dataGridView.GetRowCellValue(j, "ptname").ToString();
                                double lg = 0;
                                double.TryParse(this.dataGridView.GetRowCellValue(j, "lg").ToString(), out lg);
                                double la = 0;
                                double.TryParse(this.dataGridView.GetRowCellValue(j, "la").ToString(), out la);

                                if (dtptname == ptname)
                                {
                                    if (dtlg != lg || dtla != la)
                                    {
                                        dt.Rows[i]["选择"] = true;
                                    }
                                    else
                                    {
                                        dt.Rows[i]["选择"] = false;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
            return dt;
        }

      

       
        #region 委托事件(删除元素、纠正栅格)

        /// <summary>
        /// 从数据库替换到Excel
        /// </summary>
        /// <param name="ptlist"></param>
        void pdf_ReplaceFromDb(List<PointStruct> ptlist)
        {
            try
            {
                for (int i = 0; i < this.dataGridView.RowCount; i++)
                {
                    foreach (PointStruct ps in ptlist)
                    {
                        if (ps.ptname == this.dataGridView.GetRowCellValue(i, "ptname").ToString())
                        {
                            this.dataGridView.SetRowCellValue(i, "lg", ps.ptlg);
                            this.dataGridView.SetRowCellValue(i, "la", ps.ptla);

                            string ptname = this.dataGridView.GetRowCellValue(i, "ptname").ToString();
                            string ptlg = this.dataGridView.GetRowCellValue(i, "lg").ToString();
                            string ptla = this.dataGridView.GetRowCellValue(i, "la").ToString();
                            string dist = this.dataGridView.GetRowCellValue(i, "dist").ToString();
                            this.EditPoint.ptname = ptname;
                            double.TryParse(ptlg, out this.EditPoint.ptlg);
                            double.TryParse(ptla, out this.EditPoint.ptla);
                            this.EditPoint.dist = dist;
                            this.EditPoint.rIndex = i;

                            UpdatePtlist();
                            updateChangedPtlist(this.EditPoint);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }


          
        }

        /// <summary>
        /// 从Excel替换到数据库
        /// </summary>
        /// <param name="ptlist"></param>
        void pdf_ReplaceFromExcel(List<PointStruct> ptlist)
        {
            try
            {
                PointDao pd = new PointDao();

                for (int i = 0; i < this.dataGridView.RowCount; i++)
                {
                    foreach (PointStruct ps in ptlist)
                    {
                        if (ps.ptname == this.dataGridView.GetRowCellValue(i, "ptname").ToString())
                        {
                            string ptname = this.dataGridView.GetRowCellValue(i, "ptname").ToString();
                            string ptlg = this.dataGridView.GetRowCellValue(i, "lg").ToString();
                            string ptla = this.dataGridView.GetRowCellValue(i, "la").ToString();
                            string dist = this.dataGridView.GetRowCellValue(i, "dist").ToString();

                            double lg = 0;
                            double.TryParse(ptlg, out lg);//经度
                            double la = 0;
                            double.TryParse(ptla, out la);//经度

                            if (!double.IsNaN(lg) & !double.IsNaN(la))
                                if (lg != 0 && la != 0)
                                {
                                    try
                                    {
                                        if (ptname.Contains("'"))
                                            ptname = ptname.Replace("'", "''");
                                        //点不存在，则插入
                                        if (!pd.IsExist(ptname))
                                            pd.InserToDb(ptname, lg, la);
                                        else
                                        {
                                            pd.UpdateRow(ptname, lg, la);
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        //throw;
                                    }
                                }


                        }
                    }
                }

                pdf.UpdateDatasource(GetDtInExcel());

            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }

        /// <summary>
        /// 取中数
        /// </summary>
        /// <param name="ptlist"></param>
        void pdf_Average(List<PointStruct> ptlist)
        {
            try
            {
                PointDao pd = new PointDao();

                for (int i = 0; i < this.dataGridView.RowCount; i++)
                {
                    foreach (PointStruct ps in ptlist)
                    {
                        if (ps.ptname == this.dataGridView.GetRowCellValue(i, "ptname").ToString())
                        {
                            string ptname = this.dataGridView.GetRowCellValue(i, "ptname").ToString();
                            string ptlg = this.dataGridView.GetRowCellValue(i, "lg").ToString();
                            string ptla = this.dataGridView.GetRowCellValue(i, "la").ToString();
                            string dist = this.dataGridView.GetRowCellValue(i, "dist").ToString();

                            double lg = 0;
                            double.TryParse(ptlg, out lg);//经度
                            double la = 0;
                            double.TryParse(ptla, out la);//经度

                            if (!double.IsNaN(lg) & !double.IsNaN(la))
                                if (lg != 0 && la != 0)
                                {
                                    double avglg=(lg + ps.ptlg) / 2;
                                    double avgla=(la + ps.ptla) / 2;
                                    if (ptname.Contains("'"))
                                        ptname = ptname.Replace("'", "''");

                                    pd.UpdateRow(ptname, avglg, avgla);


                                    this.dataGridView.SetRowCellValue(i, "lg", avglg);
                                    this.dataGridView.SetRowCellValue(i, "la", avgla);
                                  
                                    this.EditPoint.ptname = ptname;
                                    this.EditPoint.ptlg = avglg;
                                    this.EditPoint.ptla = avgla;
                                    this.EditPoint.dist = dist;
                                    this.EditPoint.rIndex = i;

                                    UpdatePtlist();
                                    updateChangedPtlist(this.EditPoint);


                                }


                        }
                    }
                }

                pdf.UpdateDatasource(GetDtInExcel());

            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }
        #region 更新ptlist里的值
        private void UpdatePtlist()
        {
            List<PointStruct> pointtlist = (List<PointStruct>)resHstb[1];
            for (int j = 0; j < pointtlist.Count; j++)
                if (pointtlist[j].ptname == this.EditPoint.ptname)
                    pointtlist[j] = this.EditPoint;
        }
        #endregion
        /// <summary>
        /// 释放地图操作
        /// </summary>
        void cpf_DisposeMapOperation()
        {
            currentMapOperationType = MapOperationType.NoMode;
            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerDefault;
        }
        /// <summary>
        /// 删除所有元素
        /// </summary>
        void cpf_DeleteAllElement()
        {
            IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
            IActiveView pAcitveView = pGra as IActiveView;
            pGra.DeleteAllElements();
            pAcitveView.Refresh();
        }
        /// <summary>
        /// 删除元素
        /// </summary>
        /// <param name="index">元素name</param>
        void cpf_DeleteElement(int index)
        {
            try
            {
                IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
                IActiveView pAcitveView = pGra as IActiveView;

                pGra.Reset();
                IElement pElement = pGra.Next();

                while (pElement != null)
                {
                    IElementProperties pd = pElement as IElementProperties;
                    if (pd.Name == index.ToString())
                    {
                        pGra.DeleteElement(pElement);

                    }
                    pElement = pGra.Next();
                }

                pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                pAcitveView.Refresh();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("删除元素发生错误："+ex.Message, "错误");
            }
        }
        /// <summary>
        /// 纠正
        /// </summary>
        /// <param name="ct">控制点表</param>
        void cpf_Recify(System.Data.DataTable ct, string transTypeStr)
        {
            if (ct.Rows.Count < 3)
            {
                XtraMessageBox.Show("控制点不能少于三个！", "提示");
                return;
            }

            try
            {
                IPointCollection pFromPoint = new MultipointClass();
                IPointCollection pTPoint = new MultipointClass();

                foreach (DataRow dr in ct.Rows)
                {
                    ESRI.ArcGIS.Geometry.IPoint fpt = new PointClass() { X = double.Parse(dr[1].ToString()), Y = double.Parse(dr[2].ToString()) };
                    pFromPoint.AddPoint(fpt);
                    ESRI.ArcGIS.Geometry.IPoint tpt = new PointClass() { X = double.Parse(dr[3].ToString()), Y = double.Parse(dr[4].ToString()) };
                    pTPoint.AddPoint(tpt);
                }

                IRasterLayer rslyer = axMapControl.get_Layer(0) as IRasterLayer;
                IRaster pRaster = rslyer.Raster;

                ISpatialReferenceFactory spatialReferenceFactory = new SpatialReferenceEnvironmentClass();
                ISpatialReference spatialReference = spatialReferenceFactory.CreateGeographicCoordinateSystem((int)esriSRGeoCSType.esriSRGeoCS_WGS1984);

                switch (transTypeStr)
                {
                    case "一阶多项式":
                        GeoReferencing(pFromPoint, pTPoint, pRaster, spatialReference, esriGeoTransTypeEnum.esriGeoTransPolyOrder1);
                        break;
                    case "二阶多项式":
                        if (ct.Rows.Count >= 6)
                            GeoReferencing(pFromPoint, pTPoint, pRaster, spatialReference, esriGeoTransTypeEnum.esriGeoTransPolyOrder2);
                        else
                        {
                            XtraMessageBox.Show("二阶多项式需至少6个控制点！", "错误"); 
                            return;
                        }
                        break;
                    case "三阶多项式":
                        if (ct.Rows.Count >= 10)
                            GeoReferencing(pFromPoint, pTPoint, pRaster, spatialReference, esriGeoTransTypeEnum.esriGeoTransPolyOrder3);
                        else
                        {
                            XtraMessageBox.Show("二阶多项式需至少10个控制点！", "错误"); 
                            return;
                        }
                        break;
                    default:
                        break;
                }

               

                axMapControl.ActiveView.Refresh();

                cpf.Close();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("栅格数据配准失败，发生错误:" + ex.Message, "错误");
            }
        }

        /// <summary>
        /// 插入控制点
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        void cpf_InsertControlPoints(string x, string y,string index)
        {
            ESRI.ArcGIS.Geometry.IPoint pt = new PointClass() { X = double.Parse(x), Y = double.Parse(y) };
            IGraphicsContainer pGra = this.axMapControl.Map as IGraphicsContainer;
            IActiveView pAcitveView = pGra as IActiveView;
            IElement pEle = null;
            IMarkerElement pMakEle = new MarkerElementClass();
            pEle = pMakEle as IElement;
            ISimpleMarkerSymbol pMakSym = new SimpleMarkerSymbolClass();
            pMakSym.Style = esriSimpleMarkerStyle.esriSMSCross;
            pMakSym.Size = 7;
            pMakSym.Color = ColorToIColor(Color.Red);
            pMakEle.Symbol = pMakSym;
            pEle.Geometry = pt;

            IElementProperties pElepro = pEle as IElementProperties;
            pElepro.Name = index;

            pGra.AddElement(pEle, 0);
            pAcitveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        /// <summary>
        /// 选择控制点
        /// </summary>
        /// <param name="index"></param>
        void cpf_SelectedElement(string index)
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("删除元素发生错误：" + ex.Message, "错误");
            }
        }
        #endregion

        private void btnExit_ItemClick(object sender, ItemClickEventArgs e)
        {
            this.Close();
        }

       
        #endregion

        private void dataGridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            string lgstr = "", lastr = "";
            if (this.dataGridView.GetRowCellValue(e.RowHandle, "lg") != null)
            {
                lgstr = this.dataGridView.GetRowCellValue(e.RowHandle, "lg").ToString();
            }
            if (this.dataGridView.GetRowCellValue(e.RowHandle, "la") != null)
            {
                lastr = this.dataGridView.GetRowCellValue(e.RowHandle,"la").ToString();
            }
            if (lgstr != string.Empty && lastr != string.Empty)
            {
                if (lgstr != "0" && lastr != "0")
                {
                    e.Appearance.ForeColor = Color.Black;
                }
                else if (lgstr != "0" && lastr == "0")
                {
                    e.Appearance.ForeColor = Color.Orange;
                }
                else if (lgstr == "0" && lastr != "0")
                {
                    e.Appearance.ForeColor = Color.Orange;
                }
                else if (lgstr == "0" && lastr == "0")
                {
                    e.Appearance.ForeColor = Color.Red;
                }
            }
            else if (lgstr == string.Empty && lastr != string.Empty)
            {
                e.Appearance.ForeColor = Color.Orange;
            }
            else if (lgstr != string.Empty && lastr == string.Empty)
            {
                e.Appearance.ForeColor = Color.Orange;
            }
            else if (lgstr == string.Empty && lastr == string.Empty)
            {
                e.Appearance.ForeColor = Color.Red;
            }
        }

        private void dataGridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }  
        }

        private void repositoryItemButtonEdit1_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {

            string ptname = this.dataGridView.GetRowCellValue(dataGridView.FocusedRowHandle, "ptname").ToString();
            string ptlg = this.dataGridView.GetRowCellValue(dataGridView.FocusedRowHandle, "lg").ToString();
            string ptla = this.dataGridView.GetRowCellValue(dataGridView.FocusedRowHandle, "la").ToString();
            string dist = this.dataGridView.GetRowCellValue(dataGridView.FocusedRowHandle, "dist").ToString();

            this.axMapControl.CurrentTool = null;
            this.axMapControl.MousePointer = esriControlsMousePointer.esriPointerPencil;
            currentMapOperationType = MapOperationType.Correct;

            this.EditPoint.ptname = ptname;
            double.TryParse(ptlg, out this.EditPoint.ptlg);
            double.TryParse(ptla, out this.EditPoint.ptla);
            this.EditPoint.dist = dist;



            this.EditPoint.rIndex = dataGridView.FocusedRowHandle;
        }

        private void dataGridView_MouseDown(object sender, MouseEventArgs e)
        {

            try
            {
                DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo hInfo = dataGridView.CalcHitInfo(new System.Drawing.Point(e.X, e.Y));

                if (e.Button == MouseButtons.Left && e.Clicks == 1)
                {
                    //判断光标是否在行范围内 
                    if (hInfo.InRow && hInfo.Column.FieldName != "correct")
                    {
                        string ptname = this.dataGridView.GetRowCellValue(hInfo.RowHandle, "ptname").ToString();
                        string ptlg = this.dataGridView.GetRowCellValue(hInfo.RowHandle, "lg").ToString();
                        string ptla = this.dataGridView.GetRowCellValue(hInfo.RowHandle, "la").ToString();
                        string dist = this.dataGridView.GetRowCellValue(hInfo.RowHandle, "dist").ToString();


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

            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            try
            {
                string ptname = this.dataGridView.GetRowCellValue(e.RowHandle, "ptname").ToString();
                string ptlg = this.dataGridView.GetRowCellValue(e.RowHandle, "lg").ToString();
                string ptla = this.dataGridView.GetRowCellValue(e.RowHandle, "la").ToString();
                string dist = this.dataGridView.GetRowCellValue(e.RowHandle, "dist").ToString();

                this.EditPoint.ptname = ptname;
                double.TryParse(ptlg, out this.EditPoint.ptlg);
                double.TryParse(ptla, out this.EditPoint.ptla);
                this.EditPoint.dist = dist;
                this.EditPoint.rIndex = e.RowHandle;

                UpdatePtlist();
                updateChangedPtlist(this.EditPoint);

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("发生错误：" + ex.Message, "错误");
            }
        }

        private void btnReDraw_ItemClick(object sender, ItemClickEventArgs e)
        {
            DrawLine(false);
        }


        private string textfile = "";

        /// <summary>
        /// 拼环，苏广利
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPinhuan_ItemClick(object sender, ItemClickEventArgs e)
        {
            string linedata = string.Empty;//行数据
            string linefile = string.Empty;//线文件
            string pointfile = string.Empty;//点文件
            string cirlefile = string.Empty;//环文件
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string[] files = ofd.FileNames;

                textfile = System.IO.Path.GetDirectoryName(ofd.FileName);

                foreach (string fl in ofd.FileNames)
                {
                    if (fl.Contains("点文件.txt"))
                        pointfile = fl;
                    if (fl.Contains("线文件.txt"))
                        linefile = fl;
                    if (fl.Contains("环文件.txt"))
                        cirlefile = fl;

                }

                #region 绘制点和测段

                System.IO.StreamReader sr = new System.IO.StreamReader(linefile, Encoding.Default);
                List<string> ptlist = new List<string>();
                while ((linedata = sr.ReadLine()) != null)
                {
                    string[] linesplit = linedata.Split(',');
                    /*起始点*/
                    string frompt = linesplit[0];
                    string fromptX = GetPtlocation(pointfile, frompt)[0];
                    string fromptY = GetPtlocation(pointfile, frompt)[1];
                    ESRI.ArcGIS.Geometry.IPoint FromPt = new PointClass() { X = double.Parse(fromptX), Y = double.Parse(fromptY) };
                    if (!ptlist.Contains(frompt))
                    {
                        IElement pEle1 = null;
                        IMarkerElement pMakEle1 = new MarkerElementClass();
                        pEle1 = pMakEle1 as IElement;
                        IMarkerSymbol pMakSym1 = new SimpleMarkerSymbolClass();
                        pMakSym1.Size = 5;
                        pMakSym1.Color = ColorToIColor(Color.Turquoise);
                        pMakEle1.Symbol = pMakSym1;

                        pEle1.Geometry = FromPt;

                        IElementProperties pElepro = pEle1 as IElementProperties;
                        pElepro.Name = frompt;

                        //AddSelectedElementByGraphicsSubLayer("pointEleLayer", pEle1);

                        ptlist.Add(frompt);
                    }

                    /*终止点*/
                    string topt = linesplit[1];
                    string toptX = GetPtlocation(pointfile, topt)[0];
                    string toptY = GetPtlocation(pointfile, topt)[1];
                    ESRI.ArcGIS.Geometry.IPoint ToPt = new PointClass() { X = double.Parse(toptX), Y = double.Parse(toptY) };
                    if (!ptlist.Contains(topt))
                    {
                        IElement pEle2 = null;

                        IMarkerElement pMakEle2 = new MarkerElementClass();
                        pEle2 = pMakEle2 as IElement;
                        IMarkerSymbol pMakSym2 = new SimpleMarkerSymbolClass();
                        pMakSym2.Size = 5;
                        pMakSym2.Color = ColorToIColor(Color.Turquoise);
                        pMakEle2.Symbol = pMakSym2;
                        pEle2.Geometry = ToPt;

                        IElementProperties pElepro = pEle2 as IElementProperties;
                        pElepro.Name = topt;

                        //AddSelectedElementByGraphicsSubLayer("pointEleLayer", pEle2);

                        ptlist.Add(topt);
                    }

                    /*高差*/
                    string height = linesplit[2];
                    /*距离*/
                    string dis = linesplit[3];

                    ISegmentCollection pPath = new PathClass();
                    object o = Type.Missing;
                    ILine2 line = new LineClass();
                    line.FromPoint = FromPt;
                    line.ToPoint = ToPt;
                    pPath.AddSegment(line as ISegment, ref o, ref o);
                    IGeometryCollection pPolyline = new PolylineClass();
                    pPolyline.AddGeometry(pPath as IGeometry, ref o, ref o);
                    ISimpleLineSymbol lineSymbol = new SimpleLineSymbolClass();
                    lineSymbol.Color = ColorToIColor(Color.Red);//颜色  
                    lineSymbol.Style = esriSimpleLineStyle.esriSLSSolid; //样式  
                    lineSymbol.Width = 1;
                    ILineElement pLineElement = new LineElementClass();
                    pLineElement.Symbol = lineSymbol;
                    IElement pElement = pLineElement as IElement;
                    pElement.Geometry = pPolyline as IGeometry;

                    IElementProperties pEleproline = pElement as IElementProperties;
                    pEleproline.Name = height;

                    AddSelectedElementByGraphicsSubLayer("lineEleLayer", pElement);
                    AddArrowElement(pPolyline as IGeometry);


                }

                sr.Close();//关闭文件读取流

                #endregion

                #region 绘制环

                System.IO.StreamReader srcircle = new System.IO.StreamReader(cirlefile, Encoding.UTF8);
                string circleText=string.Empty;
                int n = 0;
              
                
                bool isReadingCirlce = false;//是否开始读取一个环
               
                while ((circleText = srcircle.ReadLine()) != null)
                {
                    n++;

                    if (n > 1)
                    {
                        /*
                         * 解析环
                         */
                        if (isReadingCirlce)
                        {
                            IPolyline2 pLine = new PolylineClass();
                            if (circleText.Contains("->"))
                            {
                                string[] pts = circleText.Split(new string[] { "->" }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (string pt in pts)
                                {
                                    string x = GetPtlocation(pointfile, pt)[0];
                                    string y = GetPtlocation(pointfile, pt)[1];
                                    ESRI.ArcGIS.Geometry.IPoint IPt = new PointClass() { X = double.Parse(x), Y = double.Parse(y) };
                                    (pLine as IPointCollection).AddPoint(IPt);

                                }
                            }
                           
                            string bhcstr = srcircle.ReadLine().Split(new string[] { "闭合差:" }, StringSplitOptions.RemoveEmptyEntries)[0];
                            double bhc = Math.Round(double.Parse(bhcstr) * 1000, 1);
                            string xccstr = srcircle.ReadLine().Split(new string[] { "限差：" }, StringSplitOptions.RemoveEmptyEntries)[0];
                            double xc = Math.Round(double.Parse(xccstr) * 1000, 1);
                           
                            ILineElement pLineElement = new LineElementClass();
                            IElement pElement;
                            pElement = pLineElement as IElement;
                            pElement.Geometry = pLine;

                            IElementProperties pEleproline = pElement as IElementProperties;
                            pEleproline.Name = "闭合差:"+bhc.ToString() + "\n" +"限差:"+ xc.ToString();

                            AddSelectedElementByGraphicsSubLayer("circleEleLayer", pElement);

                            isReadingCirlce = false;
                        }

                        //开始标记
                        string str1=circleText.Substring(0, 32);
                        string str2 = circleText.Substring(circleText.Length - 32, 32);
                        if (str1 == "================================" && str2 == "================================")
                        {
                            isReadingCirlce = true;
                        }
                        
                    }
                }
                #endregion

                axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
            }
        }

        void AddArrowElement(IGeometry pGeometry)
        {
            IRgbColor pColor = new RgbColorClass();
            pColor.Red = 255;

            ICartographicLineSymbol pCartoLineSymbol = new CartographicLineSymbolClass();
            pCartoLineSymbol.Cap = esriLineCapStyle.esriLCSRound;

            ILineProperties pLineProp = pCartoLineSymbol as ILineProperties;
            pLineProp.DecorationOnTop = true;

            ILineDecoration pLineDecoration = new LineDecorationClass();
            ISimpleLineDecorationElement pSimpleLineDecoElem = new SimpleLineDecorationElementClass();
            pSimpleLineDecoElem.AddPosition(1);
            IArrowMarkerSymbol pArrowMarkerSym = new ArrowMarkerSymbolClass();
            pArrowMarkerSym.Size = 8;
            pArrowMarkerSym.Color = pColor;
            pSimpleLineDecoElem.MarkerSymbol = pArrowMarkerSym as IMarkerSymbol;
            pLineDecoration.AddElement(pSimpleLineDecoElem as ILineDecorationElement);
            pLineProp.LineDecoration = pLineDecoration;

            ILineSymbol pLineSymbol = pCartoLineSymbol as ILineSymbol;

            pLineSymbol.Color = pColor;
            pLineSymbol.Width = 1;

            ILineElement pLineElem = new LineElementClass();
            pLineElem.Symbol = pLineSymbol;
            IElement pElem = pLineElem as IElement;
            pElem.Geometry = pGeometry;

            AddSelectedElementByGraphicsSubLayer("lineEleLayer", pElem);
        }
        /// <summary>
        /// 获取点坐标
        /// </summary>
        /// <param name="ptfile">点文件路径</param>
        /// <param name="ptname">点名</param>
        /// <returns></returns>
        private string[] GetPtlocation(string ptfile, string ptname)
        {
            string[] locations = new string[2];

            string linedata = string.Empty;
            System.IO.StreamReader sr = new System.IO.StreamReader(ptfile, Encoding.Default);
            while ((linedata = sr.ReadLine()) != null)
            {
                string[] linesplit = linedata.Split(',');
                for (int i = 0; i < linesplit.Length; i++)
                {
                    if (linesplit[0] == ptname)
                    {
                        locations[0] = linesplit[1];
                        locations[1] = linesplit[2];
                    }
                }
                
            }

            sr.Close();//关闭文件读取流

            return locations;
         }

        /// <summary>
        /// 添加标注
        /// </summary>
        /// <param name="pColor">颜色</param>
        /// <param name="pFont">字体</param>
        /// <param name="sybSize">大小</param>
        /// <param name="pEnv">显示的位置范围</param>
        /// <param name="text">显示的文本内容</param>
        private void AddText(IRgbColor pColor,IFontDisp pFont,double sybSize,IGeometry pGeo,string text)
        {
            ITextSymbol pTextSymbol = new TextSymbolClass() { Color = pColor, Font = pFont, Size = sybSize };
            ITextElement pTextElment = null;
            IElement pEle = null;
            //使用地理对象的中心作为标注的位置
            ESRI.ArcGIS.Geometry.IPoint pPoint = new PointClass();
            IEnvelope pEnv = pGeo.Envelope;
            pPoint.PutCoords(pEnv.XMin + pEnv.Width * 0.5, pEnv.YMin + pEnv.Height * 0.5);
            pTextElment = new TextElementClass() { Symbol = pTextSymbol, ScaleText = true, Text = text };
            pEle = pTextElment as IElement;
            pEle.Geometry = pPoint;

            IPointCollection pPointCollection = new PolylineClass();

            //double oldAngle = GetAngle(pPointCollection as IPolyline);
            // 获取新点角度
            //IPointCollection pointCollection = new PolylineClass();
            //pointCollection.AddPoint(centerPoint, ref missing, ref missing);
            //pointCollection.AddPoint(newPoint, ref missing, ref missing);
            // 旋转Element,角度为新旧点之差
            //ITransform2D pTransform2D = m_viewElement as ITransform2D;
            //pTransform2D.Rotate(centerPoint, (newAngle - oldAngle));

            //添加标注
            AddSelectedElementByGraphicsSubLayer("selectedEleLayer", pEle);
        }

  

        private void btnAnotPoint_ItemClick(object sender, ItemClickEventArgs e)
        {

            IGraphicsLayer sublayer = FindOrCreateGraphicsSubLayer("pointEleLayer");
            IGraphicsContainer gc = sublayer as IGraphicsContainer;
            gc.Reset();

            IElement  pElement = gc.Next();

            while (pElement != null)
            {

                IElementProperties pEleproline = pElement as IElementProperties;

                IRgbColor pColor = new RgbColorClass() { Red = 1, Blue = 1, Green = 1 };
                IFontDisp pFont = new StdFont() { Name = "宋体", Size = 5 } as IFontDisp;
                ITextSymbol pTextSymbol = new TextSymbolClass() { Color = pColor, Font = pFont, Size = 9 };
                IEnvelope pEnv = null;
                ITextElement pTextElment = null;
                IElement pEle = null;
                //使用地理对象的中心作为标注的位置
                pEnv = pElement.Geometry.Envelope;

                ESRI.ArcGIS.Geometry.IPoint pPoint = new PointClass();
                pPoint.PutCoords(pEnv.XMin + pEnv.Width * 0.5, pEnv.YMin + pEnv.Height * 0.5);

                pTextElment = new TextElementClass() { Symbol = pTextSymbol, ScaleText = true, Text = pEleproline.Name };
                pEle = pTextElment as IElement;
                pEle.Geometry = pPoint;
                //添加标注
                AddSelectedElementByGraphicsSubLayer("pointAnnotLayer", pEle);

                pElement = gc.Next();

            }
            axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        private void btnAnotLine_ItemClick(object sender, ItemClickEventArgs e)
        {
            IGraphicsLayer sublayer = FindOrCreateGraphicsSubLayer("lineEleLayer");
            IGraphicsContainer gc = sublayer as IGraphicsContainer;
            gc.Reset();

            IElement pElement = gc.Next();

            while (pElement != null)
            {

                IElementProperties pEleproline = pElement as IElementProperties;

                IRgbColor pColor = new RgbColorClass() { Red = 49, Blue = 139, Green = 87 };
                IFontDisp pFont = new StdFont() { Name = "宋体", Size = 5 } as IFontDisp;
                ITextSymbol pTextSymbol = new TextSymbolClass() { Color = pColor, Font = pFont, Size = 8 };
                IEnvelope pEnv = null;
                ITextElement pTextElment = null;
                IElement pEle = null;
                //使用地理对象的中心作为标注的位置
                pEnv = pElement.Geometry.Envelope;

                ESRI.ArcGIS.Geometry.IPoint pPoint = new PointClass();
                pPoint.PutCoords(pEnv.XMin + pEnv.Width * 0.5, pEnv.YMin + pEnv.Height * 0.5);

                pTextElment = new TextElementClass() { Symbol = pTextSymbol, ScaleText = true, Text = pEleproline.Name };
                pEle = pTextElment as IElement;
                pEle.Geometry = pPoint;
                //添加标注
                AddSelectedElementByGraphicsSubLayer("lineEleLayer", pEle);

                pElement = gc.Next();

            }
            axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }

        private void btnAnotCircle_ItemClick(object sender, ItemClickEventArgs e)
        {
            IGraphicsLayer sublayer = FindOrCreateGraphicsSubLayer("circleEleLayer");
            IGraphicsContainer gc = sublayer as IGraphicsContainer;
            gc.Reset();

            IElement pElement = gc.Next();

             FileStream fs =null;
            if (!File.Exists(textfile+"\\中心点和闭合差.txt"))
                fs = new FileStream(textfile + "\\中心点和闭合差.txt", FileMode.Create, FileAccess.Write);
            else
                fs = new FileStream(textfile + "\\中心点和闭合差.txt", FileMode.Open, FileAccess.Write);

            StreamWriter sw = new StreamWriter(fs);
            while (pElement != null)
            {

                IElementProperties pEleproline = pElement as IElementProperties;

                IRgbColor pColor = new RgbColorClass() { Red = 105, Blue = 105, Green = 105 };
                IFontDisp pFont = new StdFont() { Name = "宋体", Size = 5 } as IFontDisp;
                ITextSymbol pTextSymbol = new TextSymbolClass() { Color = pColor, Font = pFont, Size = 9 };
                IEnvelope pEnv = null;
                ITextElement pTextElment = null;
                IElement pEle = null;
                //使用地理对象的中心作为标注的位置
                pEnv = pElement.Geometry.Envelope;

                ESRI.ArcGIS.Geometry.IPoint pPoint = new PointClass();
                pPoint.PutCoords(pEnv.XMin + pEnv.Width * 0.5, pEnv.YMin + pEnv.Height * 0.5);

                pTextElment = new TextElementClass() { Symbol = pTextSymbol, ScaleText = true, Text = pEleproline.Name };
                pEle = pTextElment as IElement;
                pEle.Geometry = pPoint;
                //添加标注
                AddSelectedElementByGraphicsSubLayer("circleEleLayer", pEle);


                sw.WriteLine((pEnv.XMin + pEnv.Width * 0.5).ToString() + "," + (pEnv.YMin + pEnv.Height * 0.5).ToString() + "," + pEleproline.Name);

                pElement = gc.Next();

            }

            sw.Flush();
            sw.Close();
            fs.Close();

            axMapControl.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
        }


    }

    /// <summary>
    /// 环结构体
    /// </summary>
    public struct CircleStruct
    {
        public string CircleName ;//点名
        public bool BeginACircle ;//是否开始读取了一个环
        public List<string> Pointlist;//点集
    }
    public class PointDao
    {

        /// <summary>
        /// 判断记录是否存在
        /// </summary>
        /// <param name="ptname">水准点名</param>
        /// <returns>bool</returns>
        public bool IsExist(string ptname)
        {
            bool isexist = false;

            System.Data.DataTable dt = null;
            try 
            {
                string sql ="select * from 水准点表 where 点名= '"+ptname+"'";
                dt = AccessHelper.DataTable(sql);

                if (dt.Rows.Count > 0)
                    isexist = true;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return isexist;
        }

        /// <summary>
        /// 插入记录
        /// </summary>
        /// <param name="ptanme">点名</param>
        /// <param name="lg">经度</param>
        /// <param name="la">纬度</param>
        /// <returns>bool</returns>
        public bool InserToDb(string ptanme,double lg,double la)
        {
            bool isSuccess=false;
            try
            {
                string insertsql = "insert into 水准点表(点名,经度,纬度) values('" + ptanme + "'," + lg + "," + la + ")";
                isSuccess = AccessHelper.ExecuteSql(insertsql) > 0 ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return isSuccess;
        }


        public bool UpdateRow(string ptname, double lg, double la)
        {
            bool isSuccess = false;
            try
            {
                string insertsql = "update 水准点表 set 经度=" + lg + " , 纬度 =" + la + " where 点名='" + ptname + "'";
                isSuccess = AccessHelper.ExecuteSql(insertsql) > 0 ? true : false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return isSuccess;
        }

        public System.Data.DataTable GetPtlistDb(List<string> ptnamelist)
        {
            System.Data.DataTable dt = null;
            try
            {
                string ptnameliststr = "";
                foreach (string ptname in ptnamelist)
                {
                    ptnameliststr += "'" + ptname + "',";
                }
                ptnameliststr = ptnameliststr.Substring(0, ptnameliststr.Length - 1);

                string sql = "select * from 水准点表 where 点名 in (" + ptnameliststr + ")";
                dt = AccessHelper.DataTable(sql);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return dt;
        }




    }
}