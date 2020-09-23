namespace SDV_Outlook_Tools
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.rb_SDVOutlookTools = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.dp_Mailstatus = this.Factory.CreateRibbonDropDown();
            this.dp_Mailalter = this.Factory.CreateRibbonDropDown();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btn_MoveAttachments = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btn_ReMoveAttachments = this.Factory.CreateRibbonButton();
            this.rb_SDVOutlookTools.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rb_SDVOutlookTools
            // 
            this.rb_SDVOutlookTools.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.rb_SDVOutlookTools.Groups.Add(this.group1);
            this.rb_SDVOutlookTools.Label = "SDV-Outlook-Tools";
            this.rb_SDVOutlookTools.Name = "rb_SDVOutlookTools";
            // 
            // group1
            // 
            this.group1.Items.Add(this.dp_Mailstatus);
            this.group1.Items.Add(this.dp_Mailalter);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btn_MoveAttachments);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.btn_ReMoveAttachments);
            this.group1.Label = "Attachments";
            this.group1.Name = "group1";
            // 
            // dp_Mailstatus
            // 
            ribbonDropDownItemImpl1.Label = "gelesene";
            ribbonDropDownItemImpl2.Label = "ungelesene";
            this.dp_Mailstatus.Items.Add(ribbonDropDownItemImpl1);
            this.dp_Mailstatus.Items.Add(ribbonDropDownItemImpl2);
            this.dp_Mailstatus.Label = "Mailstatus";
            this.dp_Mailstatus.Name = "dp_Mailstatus";
            // 
            // dp_Mailalter
            // 
            ribbonDropDownItemImpl3.Label = "30";
            ribbonDropDownItemImpl4.Label = "60";
            ribbonDropDownItemImpl5.Label = "90";
            this.dp_Mailalter.Items.Add(ribbonDropDownItemImpl3);
            this.dp_Mailalter.Items.Add(ribbonDropDownItemImpl4);
            this.dp_Mailalter.Items.Add(ribbonDropDownItemImpl5);
            this.dp_Mailalter.Label = "Zeitraum";
            this.dp_Mailalter.Name = "dp_Mailalter";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btn_MoveAttachments
            // 
            this.btn_MoveAttachments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_MoveAttachments.Image = ((System.Drawing.Image)(resources.GetObject("btn_MoveAttachments.Image")));
            this.btn_MoveAttachments.Label = "Move";
            this.btn_MoveAttachments.Name = "btn_MoveAttachments";
            this.btn_MoveAttachments.ScreenTip = "Attachments from Mails.";
            this.btn_MoveAttachments.ShowImage = true;
            this.btn_MoveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MoveAttachments_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btn_ReMoveAttachments
            // 
            this.btn_ReMoveAttachments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ReMoveAttachments.Image = ((System.Drawing.Image)(resources.GetObject("btn_ReMoveAttachments.Image")));
            this.btn_ReMoveAttachments.Label = "Remove";
            this.btn_ReMoveAttachments.Name = "btn_ReMoveAttachments";
            this.btn_ReMoveAttachments.ShowImage = true;
            this.btn_ReMoveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RemoveAttachments_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.rb_SDVOutlookTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.rb_SDVOutlookTools.ResumeLayout(false);
            this.rb_SDVOutlookTools.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab rb_SDVOutlookTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_MoveAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dp_Mailstatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dp_Mailalter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ReMoveAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
