namespace SteamRouteTool
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataGridViewCellStyle gridCellStyle = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle headerCellStyle = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle alternatingCellStyle = new System.Windows.Forms.DataGridViewCellStyle();
            this.filterPanel = new System.Windows.Forms.Panel();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.btnClearFilter = new System.Windows.Forms.Button();
            this.gridHost = new System.Windows.Forms.Panel();
            this.routeGrid = new System.Windows.Forms.DataGridView();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colBlocked = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.lblOverlay = new System.Windows.Forms.Label();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnPingRoutes = new System.Windows.Forms.Button();
            this.btnClearRules = new System.Windows.Forms.Button();
            this.btnChangeGame = new System.Windows.Forms.Button();
            this.btnAbout = new System.Windows.Forms.Button();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressStatus = new System.Windows.Forms.ToolStripProgressBar();
            this.rowMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuCopyAddress = new System.Windows.Forms.ToolStripMenuItem();
            this.menuPingRow = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.menuToggleBlock = new System.Windows.Forms.ToolStripMenuItem();
            this.filterPanel.SuspendLayout();
            this.gridHost.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.routeGrid)).BeginInit();
            this.buttonPanel.SuspendLayout();
            this.statusStrip.SuspendLayout();
            this.rowMenu.SuspendLayout();
            this.SuspendLayout();
            //
            // filterPanel
            //
            this.filterPanel.BackColor = System.Drawing.SystemColors.Window;
            this.filterPanel.Controls.Add(this.txtFilter);
            this.filterPanel.Controls.Add(this.btnClearFilter);
            this.filterPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.filterPanel.Location = new System.Drawing.Point(0, 0);
            this.filterPanel.Name = "filterPanel";
            this.filterPanel.Padding = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.filterPanel.Size = new System.Drawing.Size(454, 41);
            this.filterPanel.TabIndex = 0;
            //
            // txtFilter
            //
            this.txtFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFilter.Location = new System.Drawing.Point(10, 9);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(408, 23);
            this.txtFilter.TabIndex = 0;
            this.txtFilter.TextChanged += new System.EventHandler(this.TxtFilter_TextChanged);
            //
            // btnClearFilter
            //
            this.btnClearFilter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClearFilter.FlatAppearance.BorderSize = 0;
            this.btnClearFilter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClearFilter.Location = new System.Drawing.Point(422, 9);
            this.btnClearFilter.Name = "btnClearFilter";
            this.btnClearFilter.Size = new System.Drawing.Size(22, 22);
            this.btnClearFilter.TabIndex = 1;
            this.btnClearFilter.TabStop = false;
            this.btnClearFilter.Text = "X";
            this.btnClearFilter.UseVisualStyleBackColor = true;
            this.btnClearFilter.Visible = false;
            this.btnClearFilter.Click += new System.EventHandler(this.BtnClearFilter_Click);
            //
            // gridHost
            //
            this.gridHost.Controls.Add(this.lblOverlay);
            this.gridHost.Controls.Add(this.routeGrid);
            this.gridHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridHost.Location = new System.Drawing.Point(0, 41);
            this.gridHost.Name = "gridHost";
            this.gridHost.Size = new System.Drawing.Size(454, 359);
            this.gridHost.TabIndex = 1;
            //
            // routeGrid
            //
            this.routeGrid.AllowUserToAddRows = false;
            this.routeGrid.AllowUserToDeleteRows = false;
            this.routeGrid.AllowUserToResizeColumns = false;
            this.routeGrid.AllowUserToResizeRows = false;
            alternatingCellStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.routeGrid.AlternatingRowsDefaultCellStyle = alternatingCellStyle;
            this.routeGrid.BackgroundColor = System.Drawing.SystemColors.Window;
            this.routeGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.routeGrid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.routeGrid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            headerCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            headerCellStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(247)))), ((int)(((byte)(249)))));
            headerCellStyle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(90)))), ((int)(((byte)(96)))), ((int)(((byte)(106)))));
            headerCellStyle.Padding = new System.Windows.Forms.Padding(10, 0, 0, 0);
            headerCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(247)))), ((int)(((byte)(249)))));
            headerCellStyle.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(90)))), ((int)(((byte)(96)))), ((int)(((byte)(106)))));
            this.routeGrid.ColumnHeadersDefaultCellStyle = headerCellStyle;
            this.routeGrid.ColumnHeadersHeight = 30;
            this.routeGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.routeGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colName,
            this.colPing,
            this.colBlocked});
            this.routeGrid.ContextMenuStrip = this.rowMenu;
            gridCellStyle.BackColor = System.Drawing.SystemColors.Window;
            gridCellStyle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(36)))), ((int)(((byte)(42)))));
            gridCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(232)))), ((int)(((byte)(240)))), ((int)(((byte)(252)))));
            gridCellStyle.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(36)))), ((int)(((byte)(42)))));
            gridCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.routeGrid.DefaultCellStyle = gridCellStyle;
            this.routeGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.routeGrid.EnableHeadersVisualStyles = false;
            this.routeGrid.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(239)))), ((int)(((byte)(242)))));
            this.routeGrid.Location = new System.Drawing.Point(0, 0);
            this.routeGrid.MultiSelect = false;
            this.routeGrid.Name = "routeGrid";
            this.routeGrid.RowHeadersVisible = false;
            this.routeGrid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.routeGrid.RowTemplate.Height = 28;
            this.routeGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.routeGrid.ShowCellToolTips = false;
            this.routeGrid.Size = new System.Drawing.Size(454, 359);
            this.routeGrid.TabIndex = 0;
            this.routeGrid.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.RouteGrid_CellContentClick);
            this.routeGrid.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.RouteGrid_CellMouseClick);
            this.routeGrid.CellMouseMove += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.RouteGrid_CellMouseMove);
            this.routeGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.RouteGrid_CellPainting);
            this.routeGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.RouteGrid_CellValueChanged);
            this.routeGrid.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.RouteGrid_ColumnHeaderMouseClick);
            this.routeGrid.CurrentCellDirtyStateChanged += new System.EventHandler(this.RouteGrid_CurrentCellDirtyStateChanged);
            this.routeGrid.MouseLeave += new System.EventHandler(this.RouteGrid_MouseLeave);
            this.routeGrid.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RouteGrid_KeyDown);
            //
            // colName
            //
            this.colName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colName.HeaderText = "Route";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            //
            // colPing
            //
            this.colPing.HeaderText = "Ping";
            this.colPing.Name = "colPing";
            this.colPing.ReadOnly = true;
            this.colPing.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colPing.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colPing.Width = 78;
            //
            // colBlocked
            //
            this.colBlocked.HeaderText = "Blocked";
            this.colBlocked.Name = "colBlocked";
            this.colBlocked.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colBlocked.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colBlocked.ThreeState = true;
            this.colBlocked.Width = 86;
            //
            // lblOverlay
            //
            this.lblOverlay.BackColor = System.Drawing.SystemColors.Window;
            this.lblOverlay.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblOverlay.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(122)))), ((int)(((byte)(128)))), ((int)(((byte)(138)))));
            this.lblOverlay.Location = new System.Drawing.Point(0, 0);
            this.lblOverlay.Name = "lblOverlay";
            this.lblOverlay.Size = new System.Drawing.Size(454, 359);
            this.lblOverlay.TabIndex = 1;
            this.lblOverlay.Text = "Loading routes...";
            this.lblOverlay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // buttonPanel
            //
            this.buttonPanel.Controls.Add(this.btnPingRoutes);
            this.buttonPanel.Controls.Add(this.btnClearRules);
            this.buttonPanel.Controls.Add(this.btnChangeGame);
            this.buttonPanel.Controls.Add(this.btnAbout);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.buttonPanel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.buttonPanel.Location = new System.Drawing.Point(0, 400);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Padding = new System.Windows.Forms.Padding(7, 7, 7, 7);
            this.buttonPanel.Size = new System.Drawing.Size(454, 40);
            this.buttonPanel.TabIndex = 2;
            //
            // btnPingRoutes
            //
            this.btnPingRoutes.Enabled = false;
            this.btnPingRoutes.Location = new System.Drawing.Point(348, 10);
            this.btnPingRoutes.Name = "btnPingRoutes";
            this.btnPingRoutes.Size = new System.Drawing.Size(92, 25);
            this.btnPingRoutes.TabIndex = 0;
            this.btnPingRoutes.Text = "Ping Routes";
            this.btnPingRoutes.UseVisualStyleBackColor = true;
            this.btnPingRoutes.Click += new System.EventHandler(this.BtnPingRoutes_Click);
            //
            // btnClearRules
            //
            this.btnClearRules.Enabled = false;
            this.btnClearRules.Location = new System.Drawing.Point(250, 10);
            this.btnClearRules.Name = "btnClearRules";
            this.btnClearRules.Size = new System.Drawing.Size(92, 25);
            this.btnClearRules.TabIndex = 1;
            this.btnClearRules.Text = "Clear Rules";
            this.btnClearRules.UseVisualStyleBackColor = true;
            this.btnClearRules.Click += new System.EventHandler(this.BtnClearRules_Click);
            //
            // btnChangeGame
            //
            this.btnChangeGame.Location = new System.Drawing.Point(146, 10);
            this.btnChangeGame.Name = "btnChangeGame";
            this.btnChangeGame.Size = new System.Drawing.Size(98, 25);
            this.btnChangeGame.TabIndex = 2;
            this.btnChangeGame.Text = "Change Game";
            this.btnChangeGame.UseVisualStyleBackColor = true;
            this.btnChangeGame.Click += new System.EventHandler(this.BtnChangeGame_Click);
            //
            // btnAbout
            //
            this.btnAbout.Location = new System.Drawing.Point(78, 10);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(62, 25);
            this.btnAbout.TabIndex = 3;
            this.btnAbout.Text = "About";
            this.btnAbout.UseVisualStyleBackColor = true;
            this.btnAbout.Click += new System.EventHandler(this.BtnAbout_Click);
            //
            // statusStrip
            //
            this.statusStrip.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(247)))), ((int)(((byte)(249)))));
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus,
            this.progressStatus});
            this.statusStrip.Location = new System.Drawing.Point(0, 440);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(454, 24);
            this.statusStrip.SizingGrip = false;
            this.statusStrip.TabIndex = 3;
            //
            // lblStatus
            //
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(90)))), ((int)(((byte)(96)))), ((int)(((byte)(106)))));
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(339, 19);
            this.lblStatus.Spring = true;
            this.lblStatus.Text = "Starting...";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // progressStatus
            //
            this.progressStatus.Name = "progressStatus";
            this.progressStatus.Size = new System.Drawing.Size(100, 18);
            this.progressStatus.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressStatus.Visible = false;
            //
            // rowMenu
            //
            this.rowMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuCopyAddress,
            this.menuPingRow,
            this.menuSeparator,
            this.menuToggleBlock});
            this.rowMenu.Name = "rowMenu";
            this.rowMenu.Size = new System.Drawing.Size(197, 76);
            this.rowMenu.Opening += new System.ComponentModel.CancelEventHandler(this.RowMenu_Opening);
            //
            // menuCopyAddress
            //
            this.menuCopyAddress.Name = "menuCopyAddress";
            this.menuCopyAddress.Size = new System.Drawing.Size(196, 22);
            this.menuCopyAddress.Text = "Copy IP address";
            this.menuCopyAddress.Click += new System.EventHandler(this.MenuCopyAddress_Click);
            //
            // menuPingRow
            //
            this.menuPingRow.Name = "menuPingRow";
            this.menuPingRow.Size = new System.Drawing.Size(196, 22);
            this.menuPingRow.Text = "Ping";
            this.menuPingRow.Click += new System.EventHandler(this.MenuPingRow_Click);
            //
            // menuSeparator
            //
            this.menuSeparator.Name = "menuSeparator";
            this.menuSeparator.Size = new System.Drawing.Size(193, 6);
            //
            // menuToggleBlock
            //
            this.menuToggleBlock.Name = "menuToggleBlock";
            this.menuToggleBlock.Size = new System.Drawing.Size(196, 22);
            this.menuToggleBlock.Text = "Block";
            this.menuToggleBlock.Click += new System.EventHandler(this.MenuToggleBlock_Click);
            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(470, 540);
            this.Controls.Add(this.gridHost);
            this.Controls.Add(this.filterPanel);
            this.Controls.Add(this.buttonPanel);
            this.Controls.Add(this.statusStrip);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(430, 360);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SteamRouteTool";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.filterPanel.ResumeLayout(false);
            this.filterPanel.PerformLayout();
            this.gridHost.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.routeGrid)).EndInit();
            this.buttonPanel.ResumeLayout(false);
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.rowMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel filterPanel;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Button btnClearFilter;
        private System.Windows.Forms.Panel gridHost;
        private System.Windows.Forms.DataGridView routeGrid;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPing;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colBlocked;
        private System.Windows.Forms.Label lblOverlay;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;
        private System.Windows.Forms.Button btnPingRoutes;
        private System.Windows.Forms.Button btnClearRules;
        private System.Windows.Forms.Button btnChangeGame;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.ToolStripProgressBar progressStatus;
        private System.Windows.Forms.ContextMenuStrip rowMenu;
        private System.Windows.Forms.ToolStripMenuItem menuCopyAddress;
        private System.Windows.Forms.ToolStripMenuItem menuPingRow;
        private System.Windows.Forms.ToolStripSeparator menuSeparator;
        private System.Windows.Forms.ToolStripMenuItem menuToggleBlock;
    }
}
