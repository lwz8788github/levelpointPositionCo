namespace 水准点坐标恢复软件
{
    partial class PointDbForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.gridControlPt = new DevExpress.XtraGrid.GridControl();
            this.gridViewPt = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btnDbToExcel = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.btnExcelToDb = new DevExpress.XtraEditors.SimpleButton();
            this.btnAverage = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.gridControlPt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPt)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControlPt
            // 
            this.gridControlPt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControlPt.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.gridControlPt.Location = new System.Drawing.Point(1, 0);
            this.gridControlPt.MainView = this.gridViewPt;
            this.gridControlPt.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.gridControlPt.Name = "gridControlPt";
            this.gridControlPt.Size = new System.Drawing.Size(426, 458);
            this.gridControlPt.TabIndex = 0;
            this.gridControlPt.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridViewPt});
            // 
            // gridViewPt
            // 
            this.gridViewPt.GridControl = this.gridControlPt;
            this.gridViewPt.Name = "gridViewPt";
            this.gridViewPt.OptionsSelection.MultiSelect = true;
            this.gridViewPt.OptionsView.ShowGroupPanel = false;
            // 
            // btnDbToExcel
            // 
            this.btnDbToExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDbToExcel.Location = new System.Drawing.Point(162, 466);
            this.btnDbToExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnDbToExcel.Name = "btnDbToExcel";
            this.btnDbToExcel.Size = new System.Drawing.Size(100, 30);
            this.btnDbToExcel.TabIndex = 1;
            this.btnDbToExcel.Text = "替换到Excel";
            this.btnDbToExcel.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(361, 466);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(61, 30);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnExcelToDb
            // 
            this.btnExcelToDb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExcelToDb.Location = new System.Drawing.Point(39, 465);
            this.btnExcelToDb.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnExcelToDb.Name = "btnExcelToDb";
            this.btnExcelToDb.Size = new System.Drawing.Size(117, 30);
            this.btnExcelToDb.TabIndex = 3;
            this.btnExcelToDb.Text = "替换到数据库";
            this.btnExcelToDb.Click += new System.EventHandler(this.btnExcelToDb_Click);
            // 
            // btnAverage
            // 
            this.btnAverage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAverage.Location = new System.Drawing.Point(268, 466);
            this.btnAverage.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnAverage.Name = "btnAverage";
            this.btnAverage.Size = new System.Drawing.Size(87, 30);
            this.btnAverage.TabIndex = 4;
            this.btnAverage.Text = "取中数";
            this.btnAverage.Click += new System.EventHandler(this.btnAverage_Click);
            // 
            // PointDbForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 504);
            this.Controls.Add(this.btnAverage);
            this.Controls.Add(this.btnExcelToDb);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDbToExcel);
            this.Controls.Add(this.gridControlPt);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "PointDbForm";
            this.Text = "数据对比";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.gridControlPt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewPt)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControlPt;
        private DevExpress.XtraGrid.Views.Grid.GridView gridViewPt;
        private DevExpress.XtraEditors.SimpleButton btnDbToExcel;
        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.XtraEditors.SimpleButton btnExcelToDb;
        private DevExpress.XtraEditors.SimpleButton btnAverage;
    }
}