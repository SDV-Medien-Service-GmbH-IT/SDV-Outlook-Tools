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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.rb_SDVOutlookTools = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_RemoveAttachments = this.Factory.CreateRibbonButton();
            this.dp_Mailstatus = this.Factory.CreateRibbonDropDown();
            this.dp_Mailalter = this.Factory.CreateRibbonDropDown();
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
            this.group1.Items.Add(this.btn_RemoveAttachments);
            this.group1.Label = "Attachments";
            this.group1.Name = "group1";
            // 
            // btn_RemoveAttachments
            // 
            this.btn_RemoveAttachments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_RemoveAttachments.Image = ((System.Drawing.Image)(resources.GetObject("btn_RemoveAttachments.Image")));
            this.btn_RemoveAttachments.Label = "Remove Attachments";
            this.btn_RemoveAttachments.Name = "btn_RemoveAttachments";
            this.btn_RemoveAttachments.ScreenTip = "Remove Attachments from Mails.";
            this.btn_RemoveAttachments.ShowImage = true;
            this.btn_RemoveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RemoveAttachments_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RemoveAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dp_Mailstatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dp_Mailalter;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
