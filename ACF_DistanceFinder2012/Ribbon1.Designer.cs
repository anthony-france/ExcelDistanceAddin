namespace ACF_DistanceFinder2012
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
            this.btnCalculateDistances = this.Factory.CreateRibbonButton();
            // 
            // btnCalculateDistances
            // 
            this.btnCalculateDistances.Image = global::ACF_DistanceFinder2012.Properties.Resources.bmeasureoff;
            this.btnCalculateDistances.Label = "Calculate";
            this.btnCalculateDistances.Name = "btnCalculateDistances";
            this.btnCalculateDistances.ShowImage = true;
            this.btnCalculateDistances.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalculateDistances_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            // 
            // Ribbon1.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.btnCalculateDistances);
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalculateDistances;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
