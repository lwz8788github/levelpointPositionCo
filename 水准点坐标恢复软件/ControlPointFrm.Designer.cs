namespace 水准点坐标恢复软件
{
    partial class ControlPointFrm
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
            this.gridControlCps = new DevExpress.XtraGrid.GridControl();
            this.gridViewCps = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btnDelCps = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancle = new DevExpress.XtraEditors.SimpleButton();
            this.btnRecify = new DevExpress.XtraEditors.SimpleButton();
            this.comboBoxEdit = new DevExpress.XtraEditors.ComboBoxEdit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControlCps)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewCps)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.comboBoxEdit.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControlCps
            // 
            this.gridControlCps.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.gridControlCps.Location = new System.Drawing.Point(8, 13);
            this.gridControlCps.LookAndFeel.SkinName = "Office 2010 Blue";
            this.gridControlCps.MainView = this.gridViewCps;
            this.gridControlCps.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.gridControlCps.Name = "gridControlCps";
            this.gridControlCps.Size = new System.Drawing.Size(549, 199);
            this.gridControlCps.TabIndex = 0;
            this.gridControlCps.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridViewCps});
            // 
            // gridViewCps
            // 
            this.gridViewCps.GridControl = this.gridControlCps;
            this.gridViewCps.IndicatorWidth = 50;
            this.gridViewCps.Name = "gridViewCps";
            this.gridViewCps.OptionsView.ShowGroupPanel = false;
            this.gridViewCps.DoubleClick += new System.EventHandler(this.gridViewCps_DoubleClick);
            // 
            // btnDelCps
            // 
            this.btnDelCps.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.btnDelCps.Location = new System.Drawing.Point(405, 216);
            this.btnDelCps.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnDelCps.Name = "btnDelCps";
            this.btnDelCps.Size = new System.Drawing.Size(79, 35);
            this.btnDelCps.TabIndex = 1;
            this.btnDelCps.Text = "删除";
            this.btnDelCps.Click += new System.EventHandler(this.btnDelCps_Click);
            // 
            // btnCancle
            // 
            this.btnCancle.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.btnCancle.Location = new System.Drawing.Point(489, 216);
            this.btnCancle.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCancle.Name = "btnCancle";
            this.btnCancle.Size = new System.Drawing.Size(61, 35);
            this.btnCancle.TabIndex = 3;
            this.btnCancle.Text = "取消";
            this.btnCancle.Click += new System.EventHandler(this.btnCancle_Click);
            // 
            // btnRecify
            // 
            this.btnRecify.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.btnRecify.Location = new System.Drawing.Point(320, 216);
            this.btnRecify.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnRecify.Name = "btnRecify";
            this.btnRecify.Size = new System.Drawing.Size(79, 35);
            this.btnRecify.TabIndex = 4;
            this.btnRecify.Text = "纠正";
            this.btnRecify.Click += new System.EventHandler(this.btnRecify_Click);
            // 
            // comboBoxEdit
            // 
            this.comboBoxEdit.EditValue = "一阶多项式";
            this.comboBoxEdit.Location = new System.Drawing.Point(12, 222);
            this.comboBoxEdit.Name = "comboBoxEdit";
            this.comboBoxEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.comboBoxEdit.Properties.Items.AddRange(new object[] {
            "一阶多项式",
            "二阶多项式",
            "三阶多项式"});
            this.comboBoxEdit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.comboBoxEdit.Size = new System.Drawing.Size(100, 24);
            this.comboBoxEdit.TabIndex = 5;
            // 
            // ControlPointFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 253);
            this.Controls.Add(this.comboBoxEdit);
            this.Controls.Add(this.btnRecify);
            this.Controls.Add(this.btnCancle);
            this.Controls.Add(this.btnDelCps);
            this.Controls.Add(this.gridControlCps);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ControlPointFrm";
            this.Text = "栅格数据配准控制点";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ControlPointFrm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.gridControlCps)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewCps)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.comboBoxEdit.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControlCps;
        private DevExpress.XtraGrid.Views.Grid.GridView gridViewCps;
        private DevExpress.XtraEditors.SimpleButton btnDelCps;
        private DevExpress.XtraEditors.SimpleButton btnCancle;
        private DevExpress.XtraEditors.SimpleButton btnRecify;
        private DevExpress.XtraEditors.ComboBoxEdit comboBoxEdit;
    }
}