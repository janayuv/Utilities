namespace utilities
{
    partial class UtilitiesRibbon
    {
        private System.ComponentModel.IContainer components = null;

        public UtilitiesRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnInsertSequence = this.Factory.CreateRibbonButton();
            this.btnColorDuplicates = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Utilities";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInsertSequence);
            this.group1.Items.Add(this.btnColorDuplicates);
            this.group1.Label = "Operations";
            this.group1.Name = "group1";
            // 
            // btnInsertSequence
            // 
            this.btnInsertSequence.Label = "Insert Sequence Numbers";
            this.btnInsertSequence.Name = "btnInsertSequence";
            this.btnInsertSequence.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertSequence_Click);
            // 
            // btnColorDuplicates
            // 
            this.btnColorDuplicates.Label = "Color Duplicates";
            this.btnColorDuplicates.Name = "btnColorDuplicates";
            this.btnColorDuplicates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnColorDuplicates_Click);
            // 
            // UtilitiesRibbon
            // 
            this.Name = "UtilitiesRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.UtilitiesRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertSequence;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnColorDuplicates;
    }

    partial class ThisRibbonCollection
    {
        internal UtilitiesRibbon UtilitiesRibbon
        {
            get { return this.GetRibbon<UtilitiesRibbon>(); }
        }
    }
}
