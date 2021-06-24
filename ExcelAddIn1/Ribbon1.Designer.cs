namespace ExcelAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.myAddIn = this.Factory.CreateRibbonTab();
            this.HelloWorld = this.Factory.CreateRibbonGroup();
            this.createTemplateButton = this.Factory.CreateRibbonButton();
            this.myAddIn.SuspendLayout();
            this.HelloWorld.SuspendLayout();
            this.SuspendLayout();
            // 
            // myAddIn
            // 
            this.myAddIn.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.myAddIn.Groups.Add(this.HelloWorld);
            this.myAddIn.Label = "My Add-In";
            this.myAddIn.Name = "myAddIn";
            // 
            // HelloWorld
            // 
            this.HelloWorld.Items.Add(this.createTemplateButton);
            this.HelloWorld.Label = "group1";
            this.HelloWorld.Name = "HelloWorld";
            // 
            // createTemplateButton
            // 
            this.createTemplateButton.Label = "Create Template";
            this.createTemplateButton.Name = "createTemplateButton";
            this.createTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createTemplateButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.myAddIn);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.myAddIn.ResumeLayout(false);
            this.myAddIn.PerformLayout();
            this.HelloWorld.ResumeLayout(false);
            this.HelloWorld.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab myAddIn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup HelloWorld;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createTemplateButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
