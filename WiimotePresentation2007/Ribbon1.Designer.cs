namespace WiimotePresentation2007
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// デザイナー変数が必要です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.wiiRemoteGroup = this.Factory.CreateRibbonGroup();
            this.connectWiimoteButton = this.Factory.CreateRibbonToggleButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.presentationTimerEnable = this.Factory.CreateRibbonCheckBox();
            this.timerInterval = this.Factory.CreateRibbonEditBox();
            this.message = this.Factory.CreateRibbonLabel();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.tab1.SuspendLayout();
            this.wiiRemoteGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.wiiRemoteGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // wiiRemoteGroup
            // 
            this.wiiRemoteGroup.Items.Add(this.connectWiimoteButton);
            this.wiiRemoteGroup.Items.Add(this.separator1);
            this.wiiRemoteGroup.Items.Add(this.presentationTimerEnable);
            this.wiiRemoteGroup.Items.Add(this.timerInterval);
            this.wiiRemoteGroup.Items.Add(this.message);
            this.wiiRemoteGroup.Label = "Wiimote Presentation";
            this.wiiRemoteGroup.Name = "wiiRemoteGroup";
            // 
            // connectWiimoteButton
            // 
            this.connectWiimoteButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.connectWiimoteButton.Image = ((System.Drawing.Image)(resources.GetObject("connectWiimoteButton.Image")));
            this.connectWiimoteButton.Label = "Connect Wiimote";
            this.connectWiimoteButton.Name = "connectWiimoteButton";
            this.connectWiimoteButton.ShowImage = true;
            this.connectWiimoteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.connectWiimoteButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // presentationTimerEnable
            // 
            this.presentationTimerEnable.Label = "Use timer";
            this.presentationTimerEnable.Name = "presentationTimerEnable";
            // 
            // timerInterval
            // 
            this.timerInterval.Label = "Timer interval(sec)";
            this.timerInterval.Name = "timerInterval";
            this.timerInterval.Text = null;
            // 
            // message
            // 
            this.message.Label = "Wiimote does not connect";
            this.message.Name = "message";
            // 
            // timer
            // 
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.wiiRemoteGroup.ResumeLayout(false);
            this.wiiRemoteGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup wiiRemoteGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton connectWiimoteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox presentationTimerEnable;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox timerInterval;
        private Microsoft.Office.Tools.Ribbon.RibbonLabel message;
        private System.Windows.Forms.Timer timer;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
