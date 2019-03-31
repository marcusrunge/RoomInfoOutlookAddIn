namespace RoomInfoOutlookAddIn
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class OccupancyFormRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public OccupancyFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
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
        /// Erforderliche Methode für Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.occupancyComboBox = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // occupancyComboBox
            // 
            this.occupancyComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.occupancyComboBox.FormattingEnabled = true;
            this.occupancyComboBox.Location = new System.Drawing.Point(17, 3);
            this.occupancyComboBox.Name = "occupancyComboBox";
            this.occupancyComboBox.Size = new System.Drawing.Size(369, 56);
            this.occupancyComboBox.TabIndex = 0;
            this.occupancyComboBox.SelectedIndexChanged += new System.EventHandler(this.occupancyComboBox_SelectedIndexChanged);
            // 
            // OccupancyFormRegion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.occupancyComboBox);
            this.Name = "OccupancyFormRegion";
            this.Size = new System.Drawing.Size(400, 70);
            this.FormRegionShowing += new System.EventHandler(this.OccupancyFormRegion_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.OccupancyFormRegion_FormRegionClosed);
            this.ResumeLayout(false);

        }

        #endregion

        #region Vom Formularbereich-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für Designerunterstützung. 
        /// Der Inhalt dieser Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "OccupancyFormRegion";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;

        }

        #endregion

        private System.Windows.Forms.ComboBox occupancyComboBox;

        public partial class OccupancyFormRegionFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public OccupancyFormRegionFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                OccupancyFormRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.OccupancyFormRegionFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                OccupancyFormRegion form = new OccupancyFormRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal OccupancyFormRegion OccupancyFormRegion
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(OccupancyFormRegion))
                        return (OccupancyFormRegion)item;
                }
                return null;
            }
        }
    }
}
