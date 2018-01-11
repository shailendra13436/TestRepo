// SIMS.Net
// LegalCopyright (c) Capita Business Services Ltd 1984-2009

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.IO;
using SIMS.Entities;
using SIMS.Processes;

namespace SIMS.UserInterfaces
{
    /// <summary>
    /// User control for curriculum reporting (student list/class list/reg. group list)
    /// Contains the data area panel and the report definition selection panel.
    /// </summary>
    /// <author>Ganesh G</author>
    /// <date>04th September,2009</date>
    /// <Reviewer></Reviewer>
    /// <ReviewDate></ReviewDate>
    /// <status>Un-Reviewed</status>
    /// <revision>
    /// <author>Sachin B</author>
    /// <date>13-NOV-2009</date>
    /// <details>Resolved the ambiguity between SIMS.Entities.Region and System.Drawing.Region (due to adding the referance of LookupEntities project required for pre-defined reports.) </details>    
    /// </revision>
    /// <revision>
    ///     <Author>Chandra S Neelam</Author>
    ///     <Date>09 Aug 2010</Date>
    ///     <Details>KB101227 Fix - "All Periods" report column not picking data correctly.</Details>
    /// </revision>    
    /// <Revision>
    /// <author>Arun Tiwari</author>
    /// <date>12 Sep 2013</date>
    /// <Purpose>PRB5291 fix: Set selectedIndex to zero for comboboxes DataArea and Category, when the category code is “Basic Details”. </Purpose>
    /// </Revision>
    public class ReportControl : UserControl
    {

        #region Constants
        /// <summary>
        /// String constant for panel background image
        /// </summary>
        private const string PANEL_REPORT_DEFINITION_BACKGROUND_IMAGE = "columnDefinitionAreaPanelBg";
        /// <summary>
        /// String constant for image format
        /// </summary>
        private const string IMAGE_FORMAT = ".png";
        /// <summary>
        /// String constant for the resource stream
        /// </summary>
        private const string RESOURCE_STREAM_PATH = "SIMS.UserInterfaces.Resources.";
        /// <summary>
        /// String constant for the context menu remove column
        /// </summary>
        private const string REMOVE_COLUMN = "Remove Column";
        /// <summary>
        /// String constant for the context menu rename column
        /// </summary>
        private const string RENAME_COLUMN = "Rename Column";
        /// <summary>
        /// String constant for the context menu autosize column
        /// </summary>
        private const string AUTOSIZE_COLUMN = "AutoSize Column";
        /// <summary>
        /// String constant for the context menu nest column
        /// </summary>
        private const string NEST_COLUMN = "Nest Column";
        /// <summary>
        /// String constant for the context menu remove column nesting
        /// </summary>
        private const string REMOVE_COLUMN_NESTING = "Remove Column Nesting";
        /// <summary>
        /// String constant for the context menu for student details
        /// </summary>
        private const string STUDENT_CLASSES = "Class Details";
        /// <summary>
        /// String constant for the context menu for student timetable
        /// </summary>
        private const string STUDENT_TIMETABLE = "Timetable Details";
        /// <summary>
        /// String constant for the fixed columns maximum limit exceeding error message
        /// </summary>
        private const string FIXED_COLUMNS_MESSAGE = "Exceeded the maximum fixed columns limit of ";
        /// <summary>
        /// String constant for the message for dropping a column in the fixed columns region
        /// </summary>
        private const string COLUMN_DROP_MESSAGE = ". Can't drop the column in the fixed columns region.";
        /// <summary>
        /// Don't drop Multivalue columns to fixed columns
        /// </summary>
        private const string MULTIVALUE_MESSAGE = "You can't drop multi value columns to fixed columns area.";
        /// <summary>
        /// String constant for student total
        /// </summary>        
        private const string STUDENT_TOTAL_MESSAGE = "\t Total : ";
        /// <summary>
        /// String constant for boys total
        /// </summary>        
        private const string BOYS_TOTAL_MESSAGE = "Males : ";
        /// <summary>
        /// String constant for girls total total
        /// </summary>        
        private const string Girls_TOTAL_MESSAGE = "\t Females : ";
        /// <summary>
        /// String constant student gender female
        /// </summary>        
        private const string STUDENT_GENDER_FEMALE = "Female";
        /// <summary>
        /// String constant for student gender male
        /// </summary>        
        private const string STUDENT_GENDER_MALE = "Male";
        /// <summary>
        /// Blank category code
        /// </summary>
        private const string BLANK_CATEGORY_CODE = "Blank";

        /// <summary>
        /// Remove multicolumn remove
        /// </summary>
        private const string CONFIRM_MULTICOLUMN_REMOVE = "Do you also want to remove all the other columns for the selected type?";

        #endregion

        #region Layout Controls
        private Splitter splitterDataArea;
        private Panel panelReportDefinition;
        private Panel panelReportColumns;
        private Splitter splitterReportDefinitionArea;
        private Panel panelReportDefinitionDescription;
        private TextBox textBoxReportDefinition;
        private Panel panelReportDefinitionAreaTitle;
        private Label labelReportDefintionAreaTitle;
        private Panel panelCategory;
        private ComboBox comboBoxCategory;
        private Panel panelCategoryType;
        private ComboBox comboBoxCategoryDataArea;
        private System.ComponentModel.IContainer components = null;
        private ListView listViewReportColumns;
        private HScrollBar scrollBarDataArea;
        private UIMenuItem m_UIMenuItemRemoveColumn;
        private UIMenuItem m_UIMenuItemRenameColumn;
        private UIMenuItem m_UIMenuItemAutoSizeColumn;
        private UIMenuItem m_UIMenuItemNestColumn;
        private UIMenuItem m_UIMenuItemRemoveColumnNesting;
        private UIMenuItem m_UIMenuItemStudentTimeTable;
        private UIMenuItem m_UIMenuItemStudentClassDetails;
        private UIMenuItem m_UIMenuItemSeparator;
        private System.Windows.Forms.Panel panelContainer;
        private Panel panelDataArea;
        private Panel panelReportBottom;
        private Panel panelColumnEdge;
        #endregion

        #region  Variable Declaration
        /// <summary>
        /// Process class instance
        /// </summary>
        private SIMS.Processes.ReportDetails m_ReportDetails = null;
        /// <summary>
        /// Context menu for the report column operations
        /// </summary>
        private ContextMenu m_ContextMenuColumnOptions;
        /// <summary>
        /// context menu for the report list row operations
        /// </summary>
        private ContextMenu m_ContextMenuStudentNavigationOptions;
        /// <external/>
        /// <summary>
        /// Delegate for setting the enable state of the shortcuts for column specific operations after column selection
        /// </summary>
        public DelegateEnableSelectedColumnShortcuts SetColumnOptions;
        /// <external/>
        /// <summary>
        /// Delegate for displaying the fixed columns value depending on column operations in the control
        /// </summary>
        public DelegateDisplayFixedColumns DisplayFixedColumns;
        /// <external/>
        /// <summary>
        /// Delegate for enabling the print option
        /// </summary>
        public DelegateEnablePrintOption SetPrintOption;
        /// <external/>
        /// <summary>
        /// Delegate for setting the current print list
        /// </summary>
        public DelegateSetCurrentPrintList SetCurrentPrintList;
        /// <summary>
        /// Row Span
        /// </summary>
        private int m_RowSpan = 1;
        /// <summary>
        /// Student Report data
        /// </summary>
        private StudentReportData m_StudentReportData;

        # region Printing variables
        /// <external/>
        /// <summary>
        /// Student list to be printed
        /// </summary>
        public List<StudentReportData> PrintList { get; set; }
        /// <external/>
        /// <summary>
        /// Variable to hold the column pointer of the report being printed. 
        /// </summary>
        public int PrintListColumnPointer { get; set; }
        /// <external/>
        /// <summary>
        /// Variable to hold the row pointer of the report being printed. 
        /// </summary>
        public int PrintListRowPointer { get; set; }
        /// <external/>
        /// <summary>
        /// Variable to hold the index of the page being printed (zero based index)
        /// </summary>
        public int PrintListPageIndex { get; set; }
        /// <external/>
        /// <summary>
        /// Variable to hold the index of list being printed
        /// </summary>
        public int PrintListIndex { get; set; }
        /// <external/>
        /// <summary>
        /// Collection of all lists to be included in batch printing
        /// </summary>
        public List<ReportLookup> BatchPrintLists { get; set; }        
        /// <external/>
        /// <summary>
        /// Variable to hold the count of rows of the list which can be printed per page of the report
        /// </summary>
        public int MaximumListRowsPerPage { get; set; }
        /// <external/>
        /// <summary>
        /// Variable to state if the list heder is to be printed 
        /// </summary>
        public bool IsListHeaderIsToBePrinted = true;        
        /// <external/>
        /// <summary>
        /// Variable to state if the list footer is to be printed 
        /// </summary>
        public bool IsListFooterIsToBePrinted = false;
        #endregion

        #endregion

        #region Get-Set Properties
        /// <external/>
        /// <summary>
        /// Get-Set for the process class 
        /// </summary>
        public SIMS.Processes.ReportDetails ReportDetails
        {
            get { return this.m_ReportDetails; }
            set { this.m_ReportDetails = value; }
        }

        /// <external/>
        /// <summary>
        /// Get-Set for the selected category
        /// </summary>
        public Category SelectedCategory
        {
            get
            {
                return this.comboBoxCategory.SelectedItem as Category;
            }
        }

        /// <summary>
        /// Set Report Definition Text
        /// </summary>
        public string ReportDefinitionText
        {
            set
            {
                this.textBoxReportDefinition.Text = value;
            }
            get
            {
                return this.textBoxReportDefinition.Text;
            }
        }

        /// <external/>
        /// <summary>
        /// Property to hold the drag mode of the report list drag - drop operations
        /// </summary>
        private DragModeEnum DragMode { get; set; }

        /// <summary>
        /// Enum to indicate what rows should be displayed - All Rows, Selected Rows or "Un-selected" rows
        /// </summary>
        public ReportRowDisplayTypeEnum RowDisplayType { get; set; }

        /// <summary>
        ///  ImageList icons
        /// </summary>
        public ImageList ImageListIcons { private get; set; }
        #endregion

        #region Initializer
        /// <external/>
        /// <summary>
        /// Default Constructor        
        /// </summary>
        public ReportControl()
        {
            InitializeComponent();            
            this.PrintList = new List<StudentReportData>();
            this.RowDisplayType = ReportRowDisplayTypeEnum.All;
            this.DragMode = DragModeEnum.None;
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.splitterDataArea = new SIMS.UserInterfaces.Splitter();
            this.panelReportDefinition = new SIMS.UserInterfaces.Panel();
            this.panelReportColumns = new SIMS.UserInterfaces.Panel();
            this.listViewReportColumns = new SIMS.UserInterfaces.ListView();
            this.panelCategory = new SIMS.UserInterfaces.Panel();
            this.comboBoxCategory = new SIMS.UserInterfaces.ComboBox();
            this.panelCategoryType = new SIMS.UserInterfaces.Panel();
            this.comboBoxCategoryDataArea = new SIMS.UserInterfaces.ComboBox();
            this.panelReportDefinitionAreaTitle = new SIMS.UserInterfaces.Panel();
            this.labelReportDefintionAreaTitle = new SIMS.UserInterfaces.Label();
            this.splitterReportDefinitionArea = new SIMS.UserInterfaces.Splitter();
            this.panelReportDefinitionDescription = new SIMS.UserInterfaces.Panel();
            this.textBoxReportDefinition = new SIMS.UserInterfaces.TextBox();
            this.scrollBarDataArea = new System.Windows.Forms.HScrollBar();
            this.panelContainer = new System.Windows.Forms.Panel();
            this.panelDataArea = new SIMS.UserInterfaces.Panel();
            this.panelReportBottom = new SIMS.UserInterfaces.Panel();
            this.panelColumnEdge = new SIMS.UserInterfaces.Panel();
            this.panelReportDefinition.SuspendLayout();
            this.panelReportColumns.SuspendLayout();
            this.panelCategory.SuspendLayout();
            this.panelCategoryType.SuspendLayout();
            this.panelReportDefinitionAreaTitle.SuspendLayout();
            this.panelReportDefinitionDescription.SuspendLayout();
            this.panelContainer.SuspendLayout();
            this.panelDataArea.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitterDataArea
            // 
            this.splitterDataArea.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.splitterDataArea.Dock = System.Windows.Forms.DockStyle.Right;
            this.splitterDataArea.Location = new System.Drawing.Point(743, 0);
            this.splitterDataArea.Name = "splitterDataArea";
            this.splitterDataArea.Size = new System.Drawing.Size(4, 410);
            this.splitterDataArea.TabIndex = 4;
            this.splitterDataArea.TabStop = false;
            // 
            // panelReportDefinition
            // 
            this.panelReportDefinition.BackColor = System.Drawing.SystemColors.Window;
            this.panelReportDefinition.Controls.Add(this.panelReportColumns);
            this.panelReportDefinition.Controls.Add(this.splitterReportDefinitionArea);
            this.panelReportDefinition.Controls.Add(this.panelReportDefinitionDescription);
            this.panelReportDefinition.Dock = System.Windows.Forms.DockStyle.Right;
            this.panelReportDefinition.Location = new System.Drawing.Point(747, 0);
            this.panelReportDefinition.Name = "panelReportDefinition";
            this.panelReportDefinition.Shaded = true;
            this.panelReportDefinition.Size = new System.Drawing.Size(143, 410);
            this.panelReportDefinition.TabIndex = 3;
            // 
            // panelReportColumns
            // 
            this.panelReportColumns.BackColor = System.Drawing.Color.Transparent;
            this.panelReportColumns.Controls.Add(this.listViewReportColumns);
            this.panelReportColumns.Controls.Add(this.panelCategory);
            this.panelReportColumns.Controls.Add(this.panelCategoryType);
            this.panelReportColumns.Controls.Add(this.panelReportDefinitionAreaTitle);
            this.panelReportColumns.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReportColumns.Location = new System.Drawing.Point(0, 0);
            this.panelReportColumns.Name = "panelReportColumns";
            this.panelReportColumns.Shaded = true;
            this.panelReportColumns.Size = new System.Drawing.Size(143, 238);
            this.panelReportColumns.TabIndex = 2;
            // 
            // listViewReportColumns
            // 
            this.listViewReportColumns.Attribute = null;
            this.listViewReportColumns.AutoSort = true;
            this.listViewReportColumns.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listViewReportColumns.CausesValidation = false;
            this.listViewReportColumns.CurrentItem = -1;
            this.listViewReportColumns.Description = "";
            this.listViewReportColumns.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewReportColumns.DrawAllItems = false;
            this.listViewReportColumns.FullRowSelect = true;
            this.listViewReportColumns.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listViewReportColumns.Location = new System.Drawing.Point(0, 102);
            this.listViewReportColumns.MultiSelect = false;
            this.listViewReportColumns.Name = "listViewReportColumns";
            this.listViewReportColumns.NoEdit = false;
            this.listViewReportColumns.NoUIModify = false;
            this.listViewReportColumns.Size = new System.Drawing.Size(143, 136);
            this.listViewReportColumns.TabIndex = 5;
            this.listViewReportColumns.UseCompatibleStateImageBehavior = false;
            this.listViewReportColumns.View = System.Windows.Forms.View.Details;
            this.listViewReportColumns.MouseDown += new System.Windows.Forms.MouseEventHandler(this.listViewReportColumns_MouseDown);
            // 
            // panelCategory
            // 
            this.panelCategory.BackColor = System.Drawing.Color.Transparent;
            this.panelCategory.Controls.Add(this.comboBoxCategory);
            this.panelCategory.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelCategory.Location = new System.Drawing.Point(0, 68);
            this.panelCategory.Name = "panelCategory";
            this.panelCategory.Shaded = true;
            this.panelCategory.Size = new System.Drawing.Size(143, 34);
            this.panelCategory.TabIndex = 4;
            // 
            // comboBoxCategory
            // 
            this.comboBoxCategory.Attribute = null;
            this.comboBoxCategory.CausesValidation = false;
            this.comboBoxCategory.Description = "";
            this.comboBoxCategory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBoxCategory.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxCategory.FormattingEnabled = true;
            this.comboBoxCategory.HideInactiveItems = true;
            this.comboBoxCategory.Location = new System.Drawing.Point(0, 0);
            this.comboBoxCategory.Name = "comboBoxCategory";
            this.comboBoxCategory.ReadOnly = false;
            this.comboBoxCategory.Size = new System.Drawing.Size(143, 21);
            this.comboBoxCategory.TabIndex = 2;
            this.comboBoxCategory.SelectedIndexChanged += new System.EventHandler(this.comboBoxCategory_SelectedIndexChanged);
            // 
            // panelCategoryType
            // 
            this.panelCategoryType.BackColor = System.Drawing.Color.Transparent;
            this.panelCategoryType.Controls.Add(this.comboBoxCategoryDataArea);
            this.panelCategoryType.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelCategoryType.Location = new System.Drawing.Point(0, 34);
            this.panelCategoryType.Name = "panelCategoryType";
            this.panelCategoryType.Shaded = true;
            this.panelCategoryType.Size = new System.Drawing.Size(143, 34);
            this.panelCategoryType.TabIndex = 3;
            // 
            // comboBoxCategoryDataArea
            // 
            this.comboBoxCategoryDataArea.Attribute = null;
            this.comboBoxCategoryDataArea.BackColor = System.Drawing.SystemColors.Window;
            this.comboBoxCategoryDataArea.CausesValidation = false;
            this.comboBoxCategoryDataArea.Description = "";
            this.comboBoxCategoryDataArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBoxCategoryDataArea.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxCategoryDataArea.Enabled = false;
            this.comboBoxCategoryDataArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxCategoryDataArea.FormattingEnabled = true;
            this.comboBoxCategoryDataArea.HideInactiveItems = true;
            this.comboBoxCategoryDataArea.Location = new System.Drawing.Point(0, 0);
            this.comboBoxCategoryDataArea.Name = "comboBoxCategoryDataArea";
            this.comboBoxCategoryDataArea.ReadOnly = true;
            this.comboBoxCategoryDataArea.Size = new System.Drawing.Size(143, 21);
            this.comboBoxCategoryDataArea.TabIndex = 2;
            this.comboBoxCategoryDataArea.TabStop = false;
            this.comboBoxCategoryDataArea.SelectedIndexChanged += new System.EventHandler(this.comboBoxCategoryDataArea_SelectedIndexChanged);
            // 
            // panelReportDefinitionAreaTitle
            // 
            this.panelReportDefinitionAreaTitle.BackColor = System.Drawing.Color.Transparent;
            this.panelReportDefinitionAreaTitle.Controls.Add(this.labelReportDefintionAreaTitle);
            this.panelReportDefinitionAreaTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelReportDefinitionAreaTitle.Location = new System.Drawing.Point(0, 0);
            this.panelReportDefinitionAreaTitle.Name = "panelReportDefinitionAreaTitle";
            this.panelReportDefinitionAreaTitle.Shaded = true;
            this.panelReportDefinitionAreaTitle.Size = new System.Drawing.Size(143, 34);
            this.panelReportDefinitionAreaTitle.TabIndex = 0;
            // 
            // labelReportDefintionAreaTitle
            // 
            this.labelReportDefintionAreaTitle.AutoSize = true;
            this.labelReportDefintionAreaTitle.BackColor = System.Drawing.Color.Transparent;
            this.labelReportDefintionAreaTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelReportDefintionAreaTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelReportDefintionAreaTitle.Location = new System.Drawing.Point(0, 0);
            this.labelReportDefintionAreaTitle.Name = "labelReportDefintionAreaTitle";
            this.labelReportDefintionAreaTitle.NoAlign = false;
            this.labelReportDefintionAreaTitle.NoEdit = false;
            this.labelReportDefintionAreaTitle.NoUIModify = false;
            this.labelReportDefintionAreaTitle.OriginalText = "Select Data Area";
            this.labelReportDefintionAreaTitle.Size = new System.Drawing.Size(104, 13);
            this.labelReportDefintionAreaTitle.TabIndex = 0;
            this.labelReportDefintionAreaTitle.Text = "Select Data Area";
            // 
            // splitterReportDefinitionArea
            // 
            this.splitterReportDefinitionArea.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.splitterReportDefinitionArea.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.splitterReportDefinitionArea.Location = new System.Drawing.Point(0, 238);
            this.splitterReportDefinitionArea.Name = "splitterReportDefinitionArea";
            this.splitterReportDefinitionArea.Size = new System.Drawing.Size(143, 4);
            this.splitterReportDefinitionArea.TabIndex = 1;
            this.splitterReportDefinitionArea.TabStop = false;
            // 
            // panelReportDefinitionDescription
            // 
            this.panelReportDefinitionDescription.BackColor = System.Drawing.Color.Transparent;
            this.panelReportDefinitionDescription.Controls.Add(this.textBoxReportDefinition);
            this.panelReportDefinitionDescription.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelReportDefinitionDescription.Location = new System.Drawing.Point(0, 242);
            this.panelReportDefinitionDescription.Name = "panelReportDefinitionDescription";
            this.panelReportDefinitionDescription.Shaded = true;
            this.panelReportDefinitionDescription.Size = new System.Drawing.Size(143, 168);
            this.panelReportDefinitionDescription.TabIndex = 0;
            // 
            // textBoxReportDefinition
            // 
            this.textBoxReportDefinition.Attribute = null;
            this.textBoxReportDefinition.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBoxReportDefinition.CausesValidation = false;
            this.textBoxReportDefinition.CurrentItem = -1;
            this.textBoxReportDefinition.Description = "";
            this.textBoxReportDefinition.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxReportDefinition.Location = new System.Drawing.Point(0, 0);
            this.textBoxReportDefinition.Multiline = true;
            this.textBoxReportDefinition.Name = "textBoxReportDefinition";
            this.textBoxReportDefinition.ReadOnly = true;
            this.textBoxReportDefinition.ReadOnlyPermanently = true;
            this.textBoxReportDefinition.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxReportDefinition.Size = new System.Drawing.Size(143, 168);
            this.textBoxReportDefinition.TabIndex = 0;
            this.textBoxReportDefinition.TabStop = false;
            // 
            // scrollBarDataArea
            // 
            this.scrollBarDataArea.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.scrollBarDataArea.LargeChange = 1;
            this.scrollBarDataArea.Location = new System.Drawing.Point(0, 395);
            this.scrollBarDataArea.Maximum = 5;
            this.scrollBarDataArea.Name = "scrollBarDataArea";
            this.scrollBarDataArea.Size = new System.Drawing.Size(743, 15);
            this.scrollBarDataArea.TabIndex = 6;
            this.scrollBarDataArea.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBarDataArea_Scroll);
            // 
            // panelContainer
            // 
            this.panelContainer.BackColor = System.Drawing.Color.Transparent;
            this.panelContainer.Controls.Add(this.panelDataArea);
            this.panelContainer.Controls.Add(this.scrollBarDataArea);
            this.panelContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelContainer.Location = new System.Drawing.Point(0, 0);
            this.panelContainer.Name = "panelContainer";
            this.panelContainer.Size = new System.Drawing.Size(743, 410);
            this.panelContainer.TabIndex = 7;
            // 
            // panelDataArea
            // 
            this.panelDataArea.AllowDrop = true;
            this.panelDataArea.AutoScroll = true;
            this.panelDataArea.BackColor = System.Drawing.Color.Transparent;
            this.panelDataArea.Controls.Add(this.panelReportBottom);
            this.panelDataArea.Controls.Add(this.panelColumnEdge);
            this.panelDataArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelDataArea.Location = new System.Drawing.Point(0, 0);
            this.panelDataArea.Name = "panelDataArea";
            this.panelDataArea.Shaded = true;
            this.panelDataArea.Size = new System.Drawing.Size(743, 395);
            this.panelDataArea.TabIndex = 7;
            this.panelDataArea.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.panelDataArea_MouseWheel);
            this.panelDataArea.Paint += new System.Windows.Forms.PaintEventHandler(this.panelDataArea_Paint);
            this.panelDataArea.DragOver += new System.Windows.Forms.DragEventHandler(this.panelDataArea_DragOver);
            this.panelDataArea.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panelDataArea_MouseMove);
            this.panelDataArea.Click += new System.EventHandler(this.panelDataArea_Click);
            this.panelDataArea.DragDrop += new System.Windows.Forms.DragEventHandler(this.panelDataArea_DragDrop);
            this.panelDataArea.Scroll += new System.Windows.Forms.ScrollEventHandler(this.panelDataArea_Scroll);
            this.panelDataArea.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelDataArea_MouseDown);
            this.panelDataArea.DragLeave += new System.EventHandler(this.panelDataArea_DragLeave);
            this.panelDataArea.DragEnter += new System.Windows.Forms.DragEventHandler(this.panelDataArea_DragEnter);
            // 
            // panelReportBottom
            // 
            this.panelReportBottom.BackColor = System.Drawing.Color.Transparent;
            this.panelReportBottom.Location = new System.Drawing.Point(3, 355);
            this.panelReportBottom.Name = "panelReportBottom";
            this.panelReportBottom.Shaded = true;
            this.panelReportBottom.Size = new System.Drawing.Size(19, 21);
            this.panelReportBottom.TabIndex = 7;
            // 
            // panelColumnEdge
            // 
            this.panelColumnEdge.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.panelColumnEdge.BackColor = System.Drawing.Color.Black;
            this.panelColumnEdge.Location = new System.Drawing.Point(134, 0);
            this.panelColumnEdge.Name = "panelColumnEdge";
            this.panelColumnEdge.Shaded = true;
            this.panelColumnEdge.Size = new System.Drawing.Size(2, 382);
            this.panelColumnEdge.TabIndex = 6;
            this.panelColumnEdge.Visible = false;
            // 
            // ReportControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.panelContainer);
            this.Controls.Add(this.splitterDataArea);
            this.Controls.Add(this.panelReportDefinition);
            this.Name = "ReportControl";
            this.Size = new System.Drawing.Size(890, 410);
            this.panelReportDefinition.ResumeLayout(false);
            this.panelReportColumns.ResumeLayout(false);
            this.panelCategory.ResumeLayout(false);
            this.panelCategoryType.ResumeLayout(false);
            this.panelCategoryType.PerformLayout();
            this.panelReportDefinitionAreaTitle.ResumeLayout(false);
            this.panelReportDefinitionAreaTitle.PerformLayout();
            this.panelReportDefinitionDescription.ResumeLayout(false);
            this.panelReportDefinitionDescription.PerformLayout();
            this.panelContainer.ResumeLayout(false);
            this.panelDataArea.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        #endregion

        #region Methods
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

        /// <summary>
        /// Method to set the background of the panel
        /// </summary>
        private void setPanelBackGround()
        {
            string resourceStreamKey = RESOURCE_STREAM_PATH + PANEL_REPORT_DEFINITION_BACKGROUND_IMAGE + IMAGE_FORMAT;
            System.Reflection.Assembly assembly = GetType().Assembly;
            using (Stream stream = assembly.GetManifestResourceStream(resourceStreamKey))
            {
                if (stream != null)
                {
                    Bitmap imageBitmap = new Bitmap(stream);
                    this.panelReportDefinitionAreaTitle.BackgroundImageLayout = ImageLayout.Stretch;
                    this.panelReportDefinitionAreaTitle.BackgroundImage = imageBitmap;
                }
            }
        }

        /// <summary>
        /// Method to set the control properties
        /// </summary>
        private void setLayoutControls()
        {
            
            this.components = new System.ComponentModel.Container();
            // Set the bound of the scroll bar
            this.SetHorizontalScrollBar();
            // Set the panel background
            this.setPanelBackGround();
            // Set the context menus
            this.setContextMenu();
            // Intialise the print list attributes 
            this.initialisePrintListAttributes();
        }

        /// <summary>
        /// PRB5291 Fix:Set the selectedIndex to zero for comboboxex dataarea and category.
        /// </summary>
        public void ResetDataAreaIndex()
        {
            this.bindReportDefinitionDataArea();
        }

        /// <summary>
        /// Method to bind the controls with corresponding data
        /// </summary>
        private void bindReportDefinitionDataArea()
        {
            this.comboBoxCategoryDataArea.Items.Clear();
            this.comboBoxCategoryDataArea.Populate(this.m_ReportDetails.DataAreas);
            if (this.comboBoxCategoryDataArea.Items.Count > 0)
            {
                this.comboBoxCategoryDataArea.SelectedIndex = 0;
            }
            // Disable the combo box if there is no Assessment data area
            comboBoxCategoryDataArea.Enabled = this.comboBoxCategoryDataArea.Items.Count > 1;
        }

        /// <summary>
        /// Method to bind the categories collection to the combo box
        /// </summary>
        /// <param name="dataArea">data area</param>
        private void bindCategory(DataArea dataArea)
        {
            this.comboBoxCategory.Items.Clear();
            this.comboBoxCategory.Populate(this.m_ReportDetails.GetCategoriesByArea(dataArea));
            if (this.comboBoxCategory.Items.Count > 0)
            {
                this.comboBoxCategory.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Method to bind the report columns data to the listview
        /// </summary>
        /// <param name="category">category</param>
        private void bindReportColumnsData(Category category)
        {
            // Clear the items of list view
            this.listViewReportColumns.Items.Clear();
            this.listViewReportColumns.Columns.Add("");
            // Add the listview items
            foreach (ReportColumn column in this.m_ReportDetails.AvailableColumns)
            {
                if (column.Category.Description == category.Description && !column.HiddenAttribute.Value)
                {
                    ListViewItem listviewItemReportColumn = new ListViewItem(new string[] { column.ToString() });
                    listviewItemReportColumn.Tag = column;
                    this.listViewReportColumns.Items.Add(listviewItemReportColumn);
                }
            }
            this.listViewReportColumns.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
        }

        /// <summary>
        /// Method to fill the column header area
        /// </summary>
        /// <param name="e">painteventargs</param>
        /// <param name="rectangle">column header rectangle</param>
        private void fillRectangle(PaintEventArgs e, Rectangle rectangle)
        {
            int width = rectangle.Width;
            int height = rectangle.Height;// / 3;
            int left = rectangle.X;
            int top = rectangle.Y;
            Rectangle rectTop = new Rectangle(left, top, width, height);
            top += height;
            e.Graphics.Clip = new System.Drawing.Region(rectTop);
            e.Graphics.FillRectangle(Brushes.Gainsboro, rectTop);
            e.Graphics.ResetClip();
        }

        /// <summary>
        /// Method to get the report click location details in the layout 
        /// </summary>
        /// <param name="p">Point</param>
        /// <returns>ReportControlHitInfo</returns>
        private ReportControlHitInfo getReportDataAreaLocation(Point mouseClickLocation)
        {
            ReportControlHitInfo location = new ReportControlHitInfo();
            int rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            location.InHeaderArea = mouseClickLocation.Y < rowHeight;
            int horizontalAxisPosition = 0;
            int unscrolledCols = this.m_ReportDetails.DisplayColumns.Count - this.scrollBarDataArea.Value;
            for (int displayColumnIndex = 0; displayColumnIndex < unscrolledCols; displayColumnIndex++)
            {
                int columnIndex = (displayColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns) ? displayColumnIndex : displayColumnIndex + this.scrollBarDataArea.Value;
                ReportColumn column = null;
                if (columnIndex >= 0 && columnIndex < this.m_ReportDetails.DisplayColumns.Count)
                {
                    column = this.m_ReportDetails.DisplayColumns[columnIndex];
                    if (column.NestedUnder != null)
                    {
                        continue;
                    }
                }
                if (column != null)
                {
                    horizontalAxisPosition += column.Width;
                }
                if (horizontalAxisPosition > mouseClickLocation.X && location.Column == null)
                {
                    location.Column = column;
                    location.ColumnIndex = columnIndex;
                }
                if (horizontalAxisPosition - mouseClickLocation.X < 2 && mouseClickLocation.X - horizontalAxisPosition < 2)
                {
                    location.NearDividerAtEnd = columnIndex;
                    location.DividerPosition = horizontalAxisPosition;
                }
            }
            int verticalScroll = this.panelDataArea.VerticalScroll.Value;
            rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            int headerHeight = rowHeight;
            int rowLines = this.m_RowSpan;

            int nestingPeak = this.getNestingPeak();

            //calculate cell height
            int cellHeight = (nestingPeak + 1) * rowHeight * this.m_RowSpan;
            for (int studentIndex = 0; studentIndex < this.m_ReportDetails.DisplayStudents.Count; studentIndex++)
            {
                int verticalPosition = studentIndex * cellHeight + headerHeight - verticalScroll;
                if (verticalPosition >= headerHeight && verticalPosition < this.panelDataArea.Height)
                {
                    if (mouseClickLocation.Y >= verticalPosition && mouseClickLocation.Y < verticalPosition + cellHeight)
                    {
                        location.RowIndex = studentIndex;
                    }
                }
            }
            return location;
        }

        /// <summary>
        /// Method to set the enable state of the context menus
        /// </summary>
        private void setContextMenusEnableState()
        {
            ReportColumn selectedColumn = new ReportColumn();
            int selectedColumnIndex = this.m_ReportDetails.ReportLayout.SelectedColumnIndex;
            if (selectedColumnIndex >= 0 && selectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                selectedColumn = this.m_ReportDetails.DisplayColumns[selectedColumnIndex];
            }
            this.m_UIMenuItemNestColumn.Enabled = selectedColumnIndex != 0;
            this.m_UIMenuItemRemoveColumnNesting.Enabled = selectedColumn.NestedColumns.Count > 0;
        }

        /// <summary>
        /// Method to set the context menu on right click of student details for attendance summary.
        /// </summary>
        /// <param name="contextMenuStudentSummary"></param>
        private void setContextMenu()
        {
            ImageList imageList = CheckBoxListViewImages.CreateImageList();

            // Context menu for selected student (report list row)
            this.m_ContextMenuStudentNavigationOptions = new ContextMenu();

            // Menu item student timetable 
            this.m_UIMenuItemStudentTimeTable = new UIMenuItem();
            this.m_UIMenuItemStudentTimeTable.Index = 0;
            this.m_UIMenuItemStudentTimeTable.Text = STUDENT_TIMETABLE;
            this.m_UIMenuItemStudentTimeTable.Click += new EventHandler(this.menuItemStudentTimeTable_Click);
            this.m_ContextMenuStudentNavigationOptions.MenuItems.Add(this.m_UIMenuItemStudentTimeTable);

            // Menu item student classes
            this.m_UIMenuItemStudentClassDetails = new UIMenuItem();
            this.m_UIMenuItemStudentClassDetails.Index = 1;
            this.m_UIMenuItemStudentClassDetails.Text = STUDENT_CLASSES;            
            this.m_UIMenuItemStudentClassDetails.Click += new EventHandler(this.menuItemStudentDetails_Click);
            this.m_ContextMenuStudentNavigationOptions.MenuItems.Add(this.m_UIMenuItemStudentClassDetails);

            // Context menu for selected report list column 
            this.m_ContextMenuColumnOptions = new ContextMenu();

            // Menu item rename column
            this.m_UIMenuItemRenameColumn = new UIMenuItem();
            this.m_UIMenuItemRenameColumn.Index = 0;
            this.m_UIMenuItemRenameColumn.Text = RENAME_COLUMN;
            this.m_UIMenuItemRenameColumn.ImageList = imageList;
            this.m_UIMenuItemRenameColumn.ImageIndex = (int)ButtonImage.RenameColumn;
            this.m_UIMenuItemRenameColumn.Click += new EventHandler(menuItemRenameColumn_Click);
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemRenameColumn);

            // Menu item autosize column
            this.m_UIMenuItemAutoSizeColumn = new UIMenuItem();
            this.m_UIMenuItemAutoSizeColumn.Index = 1;
            this.m_UIMenuItemAutoSizeColumn.Text = AUTOSIZE_COLUMN;
            this.m_UIMenuItemAutoSizeColumn.ImageList = imageList;
            this.m_UIMenuItemAutoSizeColumn.ImageIndex = (int)ButtonImage.AutoSize;
            this.m_UIMenuItemAutoSizeColumn.Click += new EventHandler(menuItemAutoSizeColumn_Click);
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemAutoSizeColumn);

            // Menu item remove column
            this.m_UIMenuItemRemoveColumn = new UIMenuItem();
            this.m_UIMenuItemRemoveColumn.Index = 2;
            this.m_UIMenuItemRemoveColumn.Text = REMOVE_COLUMN;
            this.m_UIMenuItemRemoveColumn.ImageList = imageList;
            this.m_UIMenuItemRemoveColumn.ImageIndex = (int)ButtonImage.RemoveColumn;
            this.m_UIMenuItemRemoveColumn.Click += new EventHandler(menuItemRemoveColumn_Click);
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemRemoveColumn);            

            // Menu item separator
            this.m_UIMenuItemSeparator = new UIMenuItem();
            this.m_UIMenuItemSeparator.Index = 3;
            this.m_UIMenuItemSeparator.Text = "-";
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemSeparator);

            // Menu item nest column
            this.m_UIMenuItemNestColumn = new UIMenuItem();
            this.m_UIMenuItemNestColumn.Index = 4;
            this.m_UIMenuItemNestColumn.Text = NEST_COLUMN;
            this.m_UIMenuItemNestColumn.ImageList = imageList;
            this.m_UIMenuItemNestColumn.ImageIndex = (int)ButtonImage.ColumnNesting;
            this.m_UIMenuItemNestColumn.Click += new EventHandler(menuItemNestColumn_Click);
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemNestColumn);

            // Menu item remove column nesting
            this.m_UIMenuItemRemoveColumnNesting = new UIMenuItem();
            this.m_UIMenuItemRemoveColumnNesting.Index = 5;
            this.m_UIMenuItemRemoveColumnNesting.Text = REMOVE_COLUMN_NESTING;
            this.m_UIMenuItemRemoveColumnNesting.ImageList = imageList;
            this.m_UIMenuItemRemoveColumnNesting.ImageIndex = (int)ButtonImage.RemoveColumnNesting;
            this.m_UIMenuItemRemoveColumnNesting.Click += new EventHandler(menuItemRemoveColumnNesting_Click);
            this.m_ContextMenuColumnOptions.MenuItems.Add(this.m_UIMenuItemRemoveColumnNesting);
        }

        /// <external/>
        /// <summary>
        /// Method to set horizontal scroll bar
        /// </summary>
        public void SetHorizontalScrollBar()
        {
            int maxscroll = this.m_ReportDetails.DisplayColumns.Count - this.m_ReportDetails.ReportLayout.FixedColumns;
            if (maxscroll > 0)
            {
                this.scrollBarDataArea.Maximum = maxscroll;
            }
        }

        /// <external/>
        /// <summary>
        /// Method to remove the nesting inside the selected column .
        /// </summary>
        /// <param name="columnIndex">column index</param>
        public void RemoveColumnNesting(int columnIndex)
        {
            if (columnIndex >= 0 && columnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];
                foreach (ReportColumn childColumn in column.NestedColumns)
                {
                    childColumn.NestedUnderAttribute = null;
                    childColumn.NestedColumns.Clear();

                }
                column.NestedColumns.Clear();
            }

            // Check if there is nesting for atleast one of the column
            if (this.m_ReportDetails.CheckIfColumnNsetingExists() == null)
            {
                this.m_ReportDetails.ReportLayout.IsNestingApplied = false;
            }

            // Reset column selection
            this.resetColumnSelection();
        }

        /// <external/>
        /// <summary>
        /// Method to nest the selected column below the previous column .
        /// </summary>
        /// <param name="columnIndex">column index</param>
        public void NestColumn(int columnIndex)
        {
            if (columnIndex > 0 && columnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                ReportColumn childColumn = this.m_ReportDetails.DisplayColumns[columnIndex];
                ReportColumn parentColumn = this.m_ReportDetails.DisplayColumns[columnIndex - 1];

                while (parentColumn.NestedUnder != null)
                {
                    parentColumn = parentColumn.NestedUnder;
                }

                childColumn.NestedUnderAttribute = parentColumn;
                parentColumn.NestedColumns.Add(childColumn);
                foreach (ReportColumn nestedColumn in childColumn.NestedColumns)
                {
                    parentColumn.NestedColumns.Add(nestedColumn);
                    nestedColumn.NestedUnderAttribute = parentColumn;
                }
                this.m_ReportDetails.ReportLayout.IsNestingApplied = true;
            }
            // Reset column selection
            this.resetColumnSelection();
        }

        /// <external/>
        /// <summary>
        /// Method to remove the selected column
        /// </summary>        
        public void RemoveColumn()
        {
            //Delete the column
            if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex >= 0 && this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {                
                ReportColumn parentColumn = this.m_ReportDetails.DisplayColumns[this.m_ReportDetails.ReportLayout.SelectedColumnIndex];
                bool deleteAllColumnsofSameType = false;
                if (parentColumn.MultivalueColumn)
                {
                    deleteAllColumnsofSameType = MessageBox.Show(this, CONFIRM_MULTICOLUMN_REMOVE, "Confirm Column Remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                }
                if (deleteAllColumnsofSameType)
                {
                    deleteMultipleColumns(parentColumn);
                }
                else
                {
                    deleteSingleColumn(parentColumn);
                }
                this.SetHorizontalScrollBar();
                // Display the fixed column value
                if (this.DisplayFixedColumns != null)
                {
                    this.DisplayFixedColumns();
                }
                // Reset column selection
                this.resetColumnSelection();
                this.m_RowSpan = 1;
            }
            // Disable the print option if there is now column in the report list
            if (this.SetPrintOption != null)
            {
                this.SetPrintOption();
            }
        }

        /// <summary>
        /// Delete all columns of the same type
        /// </summary>
        /// <param name="selectedColumn">The column to be deleted</param>
        private void deleteMultipleColumns(ReportColumn parentColumn)
        {
            for (int colIndex = this.m_ReportDetails.DisplayColumns.Count - 1; colIndex >= 0; colIndex -- )
            {
                ReportColumn selectedColumn = this.m_ReportDetails.DisplayColumns[colIndex];
                if (selectedColumn.Category == parentColumn.Category && selectedColumn.FieldName == parentColumn.FieldName)
                {
                    deleteSingleColumn(selectedColumn);
                }
            }
        }

        /// <summary>
        /// Delete a single column
        /// </summary>
        /// <param name="selectedColumn">The column to be deleted</param>
        private void deleteSingleColumn(ReportColumn selectedColumn)
        {
            foreach (ReportColumn nestedColumn in selectedColumn.NestedColumns)
            {
                this.m_ReportDetails.DisplayColumns.Remove(nestedColumn);
            }
            this.m_ReportDetails.DisplayColumns.Remove(selectedColumn);
            if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns)
            {
                this.m_ReportDetails.ReportLayout.FixedColumns--;
            }
        }

        
        /// <external/>
        /// <summary>
        /// Method to rename the selected column
        /// </summary>
        public void RenameColumn()
        {
            if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex >= 0 && this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                ReportColumn selectedColumn = this.m_ReportDetails.DisplayColumns[this.m_ReportDetails.ReportLayout.SelectedColumnIndex];
                using (DlgRenameColumn dlgRenameColumn = new DlgRenameColumn(selectedColumn.DisplayTitle, this.ImageListIcons))
                {
                    if (dlgRenameColumn.ShowDialog() == DialogResult.OK)
                    {
                        selectedColumn.DisplayTitleAttribute.Value = dlgRenameColumn.ReportColumnDisplayName;
                    }
                }
            }
        }

        /// <summary>
        /// Add new columns
        /// </summary>
        /// <param name="location">ReportControlHitInfo</param>
        private void addNewColumns(ReportControlHitInfo location)
        {
            Cursor currentCursor = this.Cursor;
            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (this.m_ReportDetails.ReportLayout.ColumnUnderDrag is ReportColumn)
                {
                    // Set row span value
                    this.m_RowSpan = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.RowSpanAttribute.Value;

                    Category category = this.m_ReportDetails.GetCategoryByCode((this.m_ReportDetails.ReportLayout.ColumnUnderDrag as ReportColumn).Category.Code);
                    if (!category.Loaded && category.Code != BLANK_CATEGORY_CODE)
                    {
                        this.m_ReportDetails.Load(this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Category);
                    }

                    // add multivalue columns
                    if (this.m_ReportDetails.ReportLayout.ColumnUnderDrag.MultivalueColumn)
                    {
                        if (this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Category.Name.ToString().ToLower() == "timetable")  //KB101227 Fix
                        {
                            IDictionary<string, string> multiColumns = this.m_ReportDetails.AllPeriods();    //KB101227 Fix
                            foreach (KeyValuePair<string, string> kvp in multiColumns)
                            {
                                ReportColumn rColumn = new ReportColumn();
                                rColumn.CategoryAttribute = category;
                                rColumn.ColumnIndexAttribute.Value = this.m_ReportDetails.DisplayColumns.Count + 1;
                                rColumn.ColumnSpanIndexAttribute.Value = Convert.ToInt32(kvp.Key);
                                rColumn.TitleAttribute.Value = kvp.Value.ToString();
                                rColumn.DisplayTitleAttribute.Value = kvp.Value.ToString();
                                rColumn.FieldNameAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.FieldName;
                                rColumn.MultivalueColumnAttribute.Value = true;
                                rColumn.NestedUnderAttribute = null;

                                rColumn.RowSpanAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.RowSpan;
                                rColumn.WidthAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Width;
                                rColumn.KeyAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Key;

                                this.m_ReportDetails.DisplayColumns.Add(rColumn);
                                this.AutoSizeColumn(this.m_ReportDetails.DisplayColumns.Count - 1);
                            }
                        }
                        else
                        {
                            int maxmultiColumns = Convert.ToInt16(this.m_ReportDetails.ReportLayout.ColumnUnderDrag.MultiColumnValue);                            

                            // Set display columns
                            for (int i = 1; i <= maxmultiColumns; i++)
                            {
                                ReportColumn rColumn = new ReportColumn();
                                rColumn.CategoryAttribute = category;
                                rColumn.ColumnIndexAttribute.Value = this.m_ReportDetails.DisplayColumns.Count + 1;
                                rColumn.ColumnSpanIndexAttribute.Value = i;
                                rColumn.DisplayTitleAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.DisplayTitle + ' ' + i;
                                rColumn.FieldNameAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.FieldName;
                                rColumn.MultivalueColumnAttribute.Value = true;
                                rColumn.NestedUnderAttribute = null;
                                rColumn.TitleAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Title + i;
                                rColumn.RowSpanAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.RowSpan;
                                rColumn.WidthAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Width;
                                rColumn.KeyAttribute.Value = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Key;

                                this.m_ReportDetails.DisplayColumns.Add(rColumn);
                                this.AutoSizeColumn(this.m_ReportDetails.DisplayColumns.Count - 1);
                            }
                        }
                    }
                    else
                    {
                        // Get the target column index inside the report list on which the column under drag is dropped                       
                        int displayColumnIndex = location.ColumnIndex < 0 ? this.m_ReportDetails.DisplayColumns.Count : location.ColumnIndex;
                        this.m_ReportDetails.DisplayColumns.Insert(displayColumnIndex, this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Clone());
                        if (location.ColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns && location.ColumnIndex >= 0)
                        {
                            this.m_ReportDetails.ReportLayout.FixedColumns++;
                        }
                        // Display the fixed column value
                        if (this.DisplayFixedColumns != null)
                        {
                            this.DisplayFixedColumns();
                        }
                        // Reset column selection
                        this.resetColumnSelection();
                        this.AutoSizeColumn(displayColumnIndex);
                    }
                }
                if (this.SetPrintOption != null)
                {
                    this.SetPrintOption();
                }
            }
            finally
            {
                this.Cursor = currentCursor;
            }
        }

        /// <external/>
        /// <summary>
        /// Method to autosize the selected column.
        /// Sets the width of the column to the max string length and if address block, adjust the row height
        /// </summary>
        /// <param name="selectedColumnIndex">columnIndex</param>
        public void AutoSizeColumn(int selectedColumnIndex)
        {
            // Resize the width relative to the header text
            this.resizeColumnWidthRelativeToTheHeaderText(selectedColumnIndex);
            // Resize column width relative to the max length of the column data field value
            if (selectedColumnIndex >= 0 && selectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                ReportColumn selectedColumn = this.m_ReportDetails.DisplayColumns[selectedColumnIndex];
                Font dataFont = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText;
                // Set the current width as max width
                int maxWidth = selectedColumn.Width;
                if (selectedColumn.FieldName == "address_block")
                {
                    m_RowSpan = selectedColumn.RowSpanAttribute.Value = 1; // force recalculation of row heights
                }
                foreach (StudentReportData student in this.m_ReportDetails.DisplayStudents)
                {
                    // Get the value for parent column width
                    int columnWidth = TextRenderer.MeasureText(student.GetColumnData(selectedColumn), dataFont, new Size(10, 10), TextFormatFlags.NoPrefix).Width;

                      if (selectedColumn.Category.Code == "FamilyHome" && selectedColumn.FieldName == "address_block")
                    {
                        int length = student.GetColumnData(selectedColumn).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                        if (length > selectedColumn.RowSpan)
                            selectedColumn.RowSpanAttribute.Value = length;
                    }


                    // for address block field, make sure the rowspan value for the column will fit the longest address (ignore the column's default rowspan)
                    if (selectedColumn.FieldName == "address_block")
                    {
                        int rowSpan = student.GetColumnData(selectedColumn).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
                        if (rowSpan > selectedColumn.RowSpan)
                        {
                            selectedColumn.RowSpanAttribute.Value = rowSpan;
                        }
                    }

>>>>>>> .r32068
                    if (columnWidth > maxWidth)
                    {
                        maxWidth = columnWidth;
                    }
                    // Get the value for nested columns width
                    foreach (ReportColumn nestedColumn in selectedColumn.NestedColumns)
                    {

                        int nestedColumnWidth = (TextRenderer.MeasureText(student.GetColumnData(nestedColumn), dataFont, new Size(10, 10), TextFormatFlags.NoPrefix)).Width;
                        if (nestedColumnWidth > maxWidth)
                        {
                            maxWidth = nestedColumnWidth;
                        }
                    }
                }
                selectedColumn.WidthAttribute.Value = maxWidth;
            }
        }

        /// <summary>
        /// Method to autosize the selected column according to the width of the column header text length.
        /// </summary>
        /// <param name="selectedColumnIndex">columnIndex</param>
        private void resizeColumnWidthRelativeToTheHeaderText(int selectedColumnIndex)
        {
            if (selectedColumnIndex >= 0 && selectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
            {
                ReportColumn selectedColumn = this.m_ReportDetails.DisplayColumns[selectedColumnIndex];

                // Don't resize blank category column
                if (selectedColumn.Category.Code == BLANK_CATEGORY_CODE)
                    return;

                Font dataFont = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText;
                // Set the current width as max width                
                int headerWidth = (TextRenderer.MeasureText(selectedColumn.DisplayTitle, dataFont, new Size(10, 10), TextFormatFlags.NoPrefix)).Width;
                selectedColumn.WidthAttribute.Value = headerWidth;
            }
        }

        /// <summary>
        /// Method to reset the column selection and corresponding column operations controls
        /// </summary>
        private void resetColumnSelection()
        {
            // Reset the selected column index
            this.m_ReportDetails.ReportLayout.SelectedColumnIndex = -1;           
            // Set enable state of the column options shortcuts 
            if (this.SetColumnOptions != null)
            {
                this.SetColumnOptions();
            }
            // Set enable state of the column options context menu
            this.setContextMenusEnableState();
        }

        /// <summary>
        /// Method to set the report list row selection
        /// </summary>
        /// <param name="location">ReportControlHitInfo</param>
        private void setRowFocus(ReportControlHitInfo location)
        {
            if (this.m_ReportDetails.ReportType == CurrReportTypeEnum.GeneralStudentList)
            {
                StudentReportData selectedStudent = null;
                if (location.RowIndex >= 0 && location.RowIndex < this.m_ReportDetails.DisplayStudents.Count)
                {
                    selectedStudent = this.m_ReportDetails.DisplayStudents[location.RowIndex];
                }

                if (selectedStudent != null)
                {
                    selectedStudent.Selected = !selectedStudent.Selected;

                    // Reset column selection
                    this.resetColumnSelection();
                    this.Invalidate();
                }
            }
        }

        /// <external/>
        /// <summary>
        /// Method to set the visibility of the meta data (report data selection) panel
        /// </summary>
        public void ToggleReportDataSelectionPanelVisibility()
        {
            this.panelReportDefinition.Visible = !this.panelReportDefinition.Visible;
            this.splitterDataArea.Visible = !this.splitterDataArea.Visible;
        }

        /// <summary>
        /// Method to get the peak value of column nesting for any of the displayed columns
        /// </summary>        
        private int getNestingPeak()
        {
            int nestingLevelPeak = 0;
            foreach (ReportColumn column in this.m_ReportDetails.DisplayColumns)
            {
                if (column.NestedColumns.Count > nestingLevelPeak)
                {
                    nestingLevelPeak = column.NestedColumns.Count;
                }
            }
            return nestingLevelPeak;
        }

        /// <summary>
        /// Method to check if the column under drag (add new column/ relocate column) can be placed in the fixed columns region.
        /// </summary>
        /// <param name="draggedColumnTargetColumnIndex">target column index of the display list</param>
        private bool isFixedColumnRangeExceeded(int draggedColumnTargetColumnIndex)
        {
            if (draggedColumnTargetColumnIndex != -1 && draggedColumnTargetColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns)
            {
                if (this.m_ReportDetails.ReportLayout.FixedColumns + 1 > this.m_ReportDetails.ReportLayout.MaximumFixedColumns && this.m_ReportDetails.DisplayColumns.Count >= this.m_ReportDetails.ReportLayout.MaximumFixedColumns)
                {
                    return true;
                }
            }
            return false;
        }

        #region Report Printing Methods

        /// <summary>
        /// Method to get the page header height for a list report print layout page
        /// </summary>   
        /// <returns>Page header height</returns>
        private int getPageHeaderHeight(PrintPageEventArgs e ,  int verticalAxisPosition)
        {          
            // Get the height of the report list title
            verticalAxisPosition += 2 * this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            // Get the height of the optional heading text
            verticalAxisPosition = this.applyListHederCalculations(e, verticalAxisPosition, this.m_ReportDetails.PrintSettings.ListHeader.Split(' '), false);
            // Get the height of the period information
            if (this.m_ReportDetails.ReportType == CurrReportTypeEnum.ClassList && this.m_ReportDetails.PrintSettings.PrintPeriodInformation)
            {
                verticalAxisPosition = this.applyListHederCalculations(e, verticalAxisPosition, this.m_ReportDetails.PrintSettings.PeriodInformationTexts, false);
            }
            return verticalAxisPosition;
        }

        /// <summary>
        /// Method to get the page footer height for a list report print layout page
        /// </summary>   
        /// <returns>Page header height</returns>
        private int getPageFooterHeight()
        {
            return this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height * this.m_ReportDetails.ReportLayout.StudentTotalText_Rows;
        }

        /// <external/>
        /// <summary>
        /// Method to get the no of rows per page for the report
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <returns>Rows per page</returns>
        public int GetRowsPerPage(PrintPageEventArgs e , int pageTop)
        {
            int rowsPerPage = 0;
            int pageHeaderHeight = this.getPageHeaderHeight(e, pageTop);
            int pageFooterHeight = this.getPageFooterHeight();
            int pageHeight = e.PageBounds.Height - pageHeaderHeight - pageFooterHeight - e.PageSettings.Margins.Bottom;
            // Get the maximum level of nesting for any column
            int nestingPeak = this.getNestingPeak();
            // Get the row height
            int rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            // Calculate the cell height
            int cellHeight = (nestingPeak + 1) * rowHeight * this.m_RowSpan;
            int columnHeaderHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            rowsPerPage = (pageHeight - columnHeaderHeight) / cellHeight;
            return rowsPerPage;
        }

        /// <summary>
        /// Method to return the maximum of all display columns widths
        /// </summary>
        /// <returns> maximum column width </returns>
        private int getMaxColumnWidth()
        {
            int maxColumnWidth = 0;
            foreach (ReportColumn column in this.m_ReportDetails.DisplayColumns)
            {
                if (column.NestedUnder == null && column.Width > maxColumnWidth)
                {
                    maxColumnWidth = column.Width;
                }
            }
            return maxColumnWidth;
        }

        /// <summary>
        /// Method to return the printarea required for the fixed columns 
        /// </summary>
        /// <returns>Summation of all fixed columns width</returns>
        private int getFixedColumnsWidth()
        {
            int fixedColumnsPrintArea = 0;
            for (int columnIndex = 0; columnIndex < this.m_ReportDetails.DisplayColumns.Count; columnIndex++)
            {
                ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];
                // Check if the column is fixed column or not
                if (columnIndex < this.m_ReportDetails.ReportLayout.FixedColumns)
                {
                    // Check if column is not nested
                    if (column.NestedUnder == null)
                    {
                        fixedColumnsPrintArea += column.Width;
                    }
                }
            }
            return fixedColumnsPrintArea;
        }
        /// <external/>
        /// <summary>
        /// Method to setup the attributes of the student list report document
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>         
        /// <returns>columns per page</returns>
        public int GetColumnsOfCurrentPage(PrintPageEventArgs e)
        {
            int lastColumnIndex = -1;
            // Get page width            
            int pageWidth = e.PageBounds.Width - e.PageSettings.Margins.Left - e.PageSettings.Margins.Right;
            // Get the maximum of all display columns widths            
            int maxColumnWidth = getMaxColumnWidth();
            int fixedColumnsPrintArea = this.getFixedColumnsWidth();
            // Check if all columns can be printed alongside the fixed columns
            if ((fixedColumnsPrintArea + maxColumnWidth) <= pageWidth)
            {
                int printableArea = fixedColumnsPrintArea;
                int fixedColumnsCount = this.m_ReportDetails.ReportLayout.FixedColumns;
                int currentPageColumnPonter = this.PrintListColumnPointer <= fixedColumnsCount - 1 ? fixedColumnsCount : this.PrintListColumnPointer - 1;
                lastColumnIndex = currentPageColumnPonter > 0 ? currentPageColumnPonter : 0;
                for (int columnIndex = lastColumnIndex; columnIndex < this.m_ReportDetails.DisplayColumns.Count; columnIndex++)
                {
                    ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];

                    if (columnIndex >= this.PrintListColumnPointer)
                    {
                        if (column.NestedUnder != null)
                        {
                            lastColumnIndex = columnIndex;
                            continue;
                        }
                        else
                        {
                            printableArea += column.Width;
                            if (printableArea <= pageWidth)
                            {
                                lastColumnIndex = columnIndex;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
                return lastColumnIndex;
            }
            return -1;
        }

        /// <summary>
        /// Method to print the list header information
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <param name="verticalAxisPosition">verticalAxisPosition</param>
        /// <param name="tokens">tokens</param>
        /// <returns>bottom of report</returns>
        private int applyListHederCalculations(PrintPageEventArgs e, int verticalAxisPosition, string[] tokens, bool includePrinting)
        {
            Font font = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText;
            Brush brush = this.m_ReportDetails.ReportLayout.Brush_Heading;
            int horizontalAxisPosition = e.PageSettings.Margins.Left;
            int pageWidth = e.PageBounds.Width - e.PageSettings.Margins.Left - e.PageSettings.Margins.Right;
            foreach (string token in tokens)
            {
                int tokenWidth = TextRenderer.MeasureText(token, font, new Size(font.Height, font.Height), TextFormatFlags.NoPrefix).Width;
                if (tokenWidth > pageWidth)
                {
                    if (horizontalAxisPosition > e.PageSettings.Margins.Left)
                    {
                        horizontalAxisPosition = e.PageSettings.Margins.Left;
                        verticalAxisPosition += font.Height;
                    }
                    if (includePrinting)
                    {
                        e.Graphics.DrawString(token, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brush, horizontalAxisPosition, verticalAxisPosition);
                    }
                    verticalAxisPosition += font.Height;
                    continue;
                }
                if (horizontalAxisPosition + tokenWidth <= pageWidth)
                {
                    if (includePrinting)
                    {
                        e.Graphics.DrawString(token, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brush, horizontalAxisPosition, verticalAxisPosition);
                    }
                    int blankSpaceWidth = TextRenderer.MeasureText(" ", font, new Size(font.Height, font.Height), TextFormatFlags.NoPrefix).Width;
                    horizontalAxisPosition += tokenWidth + blankSpaceWidth;     
                }
                else
                {
                    horizontalAxisPosition = e.PageSettings.Margins.Left;
                    verticalAxisPosition += font.Height;
                    if (includePrinting)
                    {
                        e.Graphics.DrawString(token, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brush, horizontalAxisPosition, verticalAxisPosition);
                    }                    
                    horizontalAxisPosition += tokenWidth;
                }

            }
            // Leave a blank row separator  between list header and  period information / column header 
            verticalAxisPosition += 2 * this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            return verticalAxisPosition;
        }

        /// <summary>
        /// Method to paint the column headers of the list for the current report page being printed.
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>         
        /// <param name="lastColumIndexOfTheCurrentPage">lastColumIndexOfTheCurrentPage</param>         
        /// <returns>column header top</returns>
        private int printColumnHeaders(PrintPageEventArgs e, int lastColumIndexOfTheCurrentPage, int verticalAxisPosition)
        {
            int columnHeaderTop = 0;
            int columnHeaderHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            int horizontalAxisPosition = e.PageSettings.Margins.Left;            
            if (this.IsListHeaderIsToBePrinted)
            {

                e.Graphics.DrawString(this.m_ReportDetails.PrintSettings.ListTitle, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, this.m_ReportDetails.ReportLayout.Brush_Heading, horizontalAxisPosition, verticalAxisPosition);
                // Print empty line after title (so , the top for the next text should be the title height + height for the empty line)
                verticalAxisPosition += 2 * this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
                int pageWidth = e.PageBounds.Width - e.PageSettings.Margins.Left - e.PageSettings.Margins.Right;
                string[] listHeaderTokens = this.m_ReportDetails.PrintSettings.ListHeader.Split(new char[] { '\r', '\n', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                verticalAxisPosition = this.applyListHederCalculations(e, verticalAxisPosition, listHeaderTokens, true);
                if (this.m_ReportDetails.ReportType == CurrReportTypeEnum.ClassList && this.m_ReportDetails.PrintSettings.PrintPeriodInformation)
                {
                    verticalAxisPosition = this.applyListHederCalculations(e, verticalAxisPosition, this.m_ReportDetails.PrintSettings.PeriodInformationTexts, true);
                }
            }
            columnHeaderTop = verticalAxisPosition;
            // Print column headings
            for (int columnIndex = 0; columnIndex <= lastColumIndexOfTheCurrentPage && columnIndex < this.m_ReportDetails.DisplayColumns.Count; columnIndex++)
            {
                if (columnIndex < this.m_ReportDetails.ReportLayout.FixedColumns || columnIndex >= this.PrintListColumnPointer)
                {
                    // Get report column object
                    ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];
                    if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                    {
                        continue;
                    }
                    string text = column.DisplayTitle;
                    Rectangle rect = new Rectangle(horizontalAxisPosition, columnHeaderTop, column.Width , columnHeaderHeight);
                    e.Graphics.Clip = new System.Drawing.Region(rect);
                    Brush brush = this.m_ReportDetails.ReportLayout.Brush_Heading;
                    e.Graphics.DrawString(text, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brush, horizontalAxisPosition, columnHeaderTop);                    
                    e.Graphics.ResetClip();
                    horizontalAxisPosition += column.Width ;
                    // Set maximum rowspan value
                    if (column.RowSpanAttribute.Value > this.m_RowSpan)
                    {
                        this.m_RowSpan = column.RowSpanAttribute.Value;
                    }
                }
            }

            // Print column header border
            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThickLine, new Point(e.PageSettings.Margins.Left, columnHeaderTop), new Point(horizontalAxisPosition, columnHeaderTop));
            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThickLine, new Point(e.PageSettings.Margins.Left, columnHeaderTop + columnHeaderHeight), new Point(horizontalAxisPosition, columnHeaderTop + columnHeaderHeight));

            return columnHeaderTop;
        }

        /// <summary>
        /// Method to paint the student data of the list for the current report page being printed.
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>        
        /// <param name="lastColumIndexOfTheCurrentPage">lastColumIndexOfTheCurrentPage</param>        
        /// <param name="columnHeaderTop">columnHeaderTop</param>
        /// <returns>bottom of the page</returns>
        private int printStudentData(PrintPageEventArgs e, int lastColumIndexOfTheCurrentPage, int columnHeaderTop)
        {
            int columnHeaderHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            int rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            int nestingPeak = this.getNestingPeak();            
            int studentCount = this.PrintList.Count;
            int lastRowIndex = (this.PrintListRowPointer + this.MaximumListRowsPerPage) < studentCount ? (this.PrintListRowPointer + this.MaximumListRowsPerPage) : studentCount;
            int horizontalAxisPosition = e.PageSettings.Margins.Left;
            int dividerFrequency = this.m_ReportDetails.ReportLayout.DividerFrequency;
            int pageFooterHeight = this.getPageFooterHeight();
            int studentListTop = columnHeaderTop + columnHeaderHeight;
            int pageHeight = e.PageBounds.Height - columnHeaderTop - e.PageSettings.Margins.Bottom - pageFooterHeight;
            int verticalAxisPosition = 0;
            int bottomOfReport = columnHeaderTop + columnHeaderHeight;
            int rowIndex = 0;
            //Draw student data
            int cellHeight = (nestingPeak + 1) * rowHeight * this.m_RowSpan;
            for (int studentIndex = this.PrintListRowPointer; studentIndex < lastRowIndex; studentIndex++)
            {
                verticalAxisPosition = rowIndex * cellHeight + studentListTop;
                StudentReportData student = this.PrintList[studentIndex];
                horizontalAxisPosition = e.PageSettings.Margins.Left;
                for (int displayColumnindex = 0; displayColumnindex <= lastColumIndexOfTheCurrentPage && displayColumnindex < this.m_ReportDetails.DisplayColumns.Count; displayColumnindex++)
                {
                    if (displayColumnindex >= this.m_ReportDetails.ReportLayout.FixedColumns && displayColumnindex < this.PrintListColumnPointer)
                    {
                        continue;
                    }
                    ReportColumn column = this.m_ReportDetails.DisplayColumns[displayColumnindex];
                    if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                    {
                        continue;
                    }
                    Rectangle rect = new Rectangle(horizontalAxisPosition, verticalAxisPosition, column.Width, cellHeight);
                    e.Graphics.Clip = new System.Drawing.Region(rect);
                    ReportColumn selectedColumn = null;
                    if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex >= 0 && this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
                    {
                        selectedColumn = this.m_ReportDetails.DisplayColumns[this.m_ReportDetails.ReportLayout.SelectedColumnIndex];
                    }
                    Brush brush = this.m_ReportDetails.ReportLayout.Brush_Data;
                    // Get column Data
                    string parentFieldText = student.GetColumnData(column);
                    e.Graphics.DrawString(parentFieldText, this.m_ReportDetails.ReportLayout.Font_ReportList_DataText, brush, horizontalAxisPosition, verticalAxisPosition);
                    bottomOfReport = verticalAxisPosition + cellHeight;
                    if (this.m_ReportDetails.ReportLayout.IsNestingApplied && column.NestedColumns.Count > 0)
                    {
                        int verticalPosition = verticalAxisPosition;
                        foreach (ReportColumn nestedColumn in column.NestedColumns)
                        {
                            string nestedColumnFieldText = student.GetColumnData(nestedColumn);
                            verticalPosition += (rowHeight * this.m_RowSpan);
                            e.Graphics.DrawString(nestedColumnFieldText, this.m_ReportDetails.ReportLayout.Font_ReportList_DataText, brush, horizontalAxisPosition, verticalPosition);
                        }
                    }
                    e.Graphics.ResetClip();
                    horizontalAxisPosition += column.Width;
                    // Print dividing lines
                    if (dividerFrequency > 0)
                    {
                        if (rowIndex % dividerFrequency == 0 && rowIndex != 0)
                        {
                            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThinLine, new Point(e.PageSettings.Margins.Left, verticalAxisPosition), new Point(horizontalAxisPosition, verticalAxisPosition));
                        }
                    }
                }
                rowIndex++;
            }
            return bottomOfReport;
        }

        /// <summary>
        ///  Method to paint the student data of the list for the current report page being printed.
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <param name="bottomOfPage">bottomOfPage</param>        
        /// <param name="lastColumIndexOfTheCurrentPage">lastColumIndexOfTheCurrentPage</param>
        private int printVerticalLines(PrintPageEventArgs e, int columnHeaderTop, int bottomOfList, int lastColumIndexOfTheCurrentPage)
        {
            int rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            int pageFooterHeight = this.getPageFooterHeight();
            int pageHeight = e.PageBounds.Height - columnHeaderTop - e.PageSettings.Margins.Bottom - pageFooterHeight;
            int horizontalAxisPosition = e.PageSettings.Margins.Left;
            int fixedColumnIndex = this.m_ReportDetails.ReportLayout.FixedColumns;

            for (int displayColumnindex = 0; displayColumnindex <= lastColumIndexOfTheCurrentPage && displayColumnindex < this.m_ReportDetails.DisplayColumns.Count; displayColumnindex++)
            {
                // Dont print the line if the column is neither a fixed column nor it comes under the current columns being printed
                if (displayColumnindex >= this.m_ReportDetails.ReportLayout.FixedColumns && displayColumnindex < this.PrintListColumnPointer)
                {
                    continue;
                }
                ReportColumn column = this.m_ReportDetails.DisplayColumns[displayColumnindex];
                // Skip the vertical line divider if the column is nested under another column
                if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                {
                    fixedColumnIndex++;
                    continue;
                }
                Pen pen = (displayColumnindex == fixedColumnIndex) ? this.m_ReportDetails.ReportLayout.Pen_ColumnDividerLine : this.m_ReportDetails.ReportLayout.Pen_ThinLine;
                e.Graphics.DrawLine(pen, new Point(horizontalAxisPosition, columnHeaderTop), new Point(horizontalAxisPosition, bottomOfList));
                horizontalAxisPosition += column.Width ;
            }
            Pen drawingPen = (this.m_ReportDetails.DisplayColumns.Count == fixedColumnIndex) ? this.m_ReportDetails.ReportLayout.Pen_ColumnDividerLine : this.m_ReportDetails.ReportLayout.Pen_ThinLine;
            // Print the vertical line after last display column
            e.Graphics.DrawLine(drawingPen, new Point(horizontalAxisPosition, columnHeaderTop), new Point(horizontalAxisPosition, bottomOfList));
            // Print the horizontal line at the end of the student list
            if (this.PrintList.Count > 0)
            {
                e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThinLine, new Point(e.PageSettings.Margins.Left, bottomOfList), new Point(horizontalAxisPosition, bottomOfList));
            }
            // Print footer information 
            // Leave a blank row separator between list bottom and footer text
            if (this.IsListFooterIsToBePrinted && this.m_ReportDetails.PrintSettings.PrintTotalFigures)
            {
                int pageFooterTop = bottomOfList + rowHeight;
                string total = STUDENT_TOTAL_MESSAGE + this.PrintList.Count.ToString();
                string boysTotal = BOYS_TOTAL_MESSAGE + this.m_ReportDetails.GetStudentsByGender(this.PrintList, STUDENT_GENDER_MALE).Count.ToString();
                string girlsTotal = Girls_TOTAL_MESSAGE + this.m_ReportDetails.GetStudentsByGender(this.PrintList, STUDENT_GENDER_FEMALE).Count.ToString();
                string footerText = boysTotal + girlsTotal + total;
                Brush brushPageFooter = this.m_ReportDetails.ReportLayout.Brush_Heading;
                e.Graphics.DrawString(footerText, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brushPageFooter, e.PageSettings.Margins.Left, pageFooterTop);
                // Return pageFooterTop as the bottom of page with additional separator margin of a blank row between two lists.
                return pageFooterTop + (2 * this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height); 
            }
            // Return the bottom of page
            return bottomOfList;
          
        }
        // <external/>
        /// <summary>
        /// Method to print page no. 
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>        
        public void PrintPageBottom(PrintPageEventArgs e)
        {
            // Print page no.
            Font pageTextFont = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText;
            int pageTextTop = e.PageBounds.Height - e.PageSettings.Margins.Bottom - pageTextFont.Height;
            int pageTextLeft = (e.PageBounds.Width - e.PageSettings.Margins.Left - e.PageSettings.Margins.Right) / 2;
            e.Graphics.DrawString((this.PrintListPageIndex + 1).ToString(), pageTextFont, this.m_ReportDetails.ReportLayout.Brush_Data, pageTextLeft, pageTextTop);
            // Print date time stamp
            Font fontPageBottomText = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText;
            string dateTimeText = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            int dateTimeTextWidth = TextRenderer.MeasureText(dateTimeText, fontPageBottomText, new Size(10, 10), TextFormatFlags.NoPrefix).Width;
            int dateTimeTextTop = e.PageBounds.Height - e.PageSettings.Margins.Bottom - fontPageBottomText.Height;
            int dateTimeTextLeft = e.PageBounds.Width - e.PageSettings.Margins.Right - dateTimeTextWidth;
            if (dateTimeTextLeft >= e.PageSettings.Margins.Left)
            {
                e.Graphics.DrawString(dateTimeText, fontPageBottomText, this.m_ReportDetails.ReportLayout.Brush_Data, dateTimeTextLeft, dateTimeTextTop);
            }
        }
        /// <external/>
        /// <summary>
        /// Method to paint the report page
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <returns>lastColumIndexOfTheCurrentPage</returns>
        public void PrintReport(PrintPageEventArgs e , int verticalAxisPosition)
        {
            int lastColumIndexOfTheCurrentPage = this.GetColumnsOfCurrentPage(e);
            int columnHeaderTop = this.printColumnHeaders(e, lastColumIndexOfTheCurrentPage, verticalAxisPosition);
            int bottomOfList = this.printStudentData(e, lastColumIndexOfTheCurrentPage, columnHeaderTop);
            int bottomOfPage = this.printVerticalLines(e, columnHeaderTop, bottomOfList, lastColumIndexOfTheCurrentPage);
            // For starting the next list on the same page check if its a batch print operation , page break is not required and all lists are not yet printed.
            bool isNextListToBeStartedFromCurrentBottomOfPage = !this.m_ReportDetails.PrintSettings.PageBreakRequired && !this.m_ReportDetails.PrintSettings.PrintOnlyCurrentList && (this.PrintListIndex + 1) < this.BatchPrintLists.Count;
            if (isNextListToBeStartedFromCurrentBottomOfPage)
            {
                if (this.PrintListRowPointer + this.MaximumListRowsPerPage > this.PrintList.Count)
                {
                    this.PrintListRowPointer = 0;
                }
                // Check if another list can be started from bottomOfPage.  
                if (ValidateMargins(e, bottomOfPage))
                {
                    // Set the current list  being printed as the next list.
                    if (this.SetCurrentPrintList != null)
                    {
                        this.PrintListIndex++;
                        this.SetCurrentPrintList();
                        // Get the no rows per page.
                        this.MaximumListRowsPerPage = this.GetRowsPerPage(e, bottomOfPage);
                        this.IsListHeaderIsToBePrinted = true;
                        // Set the footer information text flag for the last page of the list (including repeated columns.)
                        this.IsListFooterIsToBePrinted = this.PrintListRowPointer + this.MaximumListRowsPerPage >= this.PrintList.Count;

                    }
                    // Call the same method recursively.
                    this.PrintReport(e, bottomOfPage);
                }
                else
                {                   
                    this.IsListHeaderIsToBePrinted = false;
                    this.IsListFooterIsToBePrinted = false;
                    return;
                }
            } 
           
        }
        /// <summary>
        /// Method to validate if the report list can be printed inside the set margins
        /// </summary>
        /// <param name="e">PrintPageEventArgs</param>
        /// <returns>lastColumIndexOfTheCurrentPage</returns>
        public bool ValidateMargins(PrintPageEventArgs e , int pageTop)
        {
            int headerFontHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            int dataFontHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            int pageHeight = e.PageBounds.Height - e.PageSettings.Margins.Top - e.PageSettings.Margins.Bottom;
            // Get the minimum height of the report page  
            int minimumReportArea = pageTop + e.PageSettings.Margins.Bottom;
            // Add the title height + blank row height 
            minimumReportArea += 2 * headerFontHeight;
            // Add the height of the optional list header text 
            minimumReportArea = this.applyListHederCalculations(e, minimumReportArea, this.m_ReportDetails.PrintSettings.ListHeader.Split(' '), false);
            if (this.m_ReportDetails.ReportType == CurrReportTypeEnum.ClassList && this.m_ReportDetails.PrintSettings.PrintPeriodInformation)
            {
                // Add the height of the period information text 
                minimumReportArea = this.applyListHederCalculations(e, minimumReportArea, this.m_ReportDetails.PrintSettings.ListHeader.Split(' '), false);
            }
            // Add the height of the column header
            minimumReportArea += headerFontHeight;
            // Add the height of single cell
            // Get the student list cell height
            int cellHeight = (this.getNestingPeak() + 1) * this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height * this.m_RowSpan;
            minimumReportArea += cellHeight;
            // Add the total figures information text height 
            minimumReportArea += 2 * dataFontHeight;
            // Add the page index  and date time text height
            minimumReportArea += 2 * dataFontHeight;
            if (pageHeight < minimumReportArea)
            {
                return false;
            }
            // Get minimum width of the page area as addition of fixed columns area and the column having max width                                            
            int minReportWidth = this.getMaxColumnWidth() + this.getFixedColumnsWidth();

            if (e.PageBounds.Width - e.PageSettings.Margins.Left - e.PageSettings.Margins.Right < minReportWidth)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// Method to initialise the print list attributes
        /// </summary>
        private void initialisePrintListAttributes()
        {
            this.PrintListColumnPointer = -1;
            this.PrintListRowPointer = 0;
            this.PrintListPageIndex = 0;
            this.PrintListIndex = 0;
            this.MaximumListRowsPerPage = 0;
            this.BatchPrintLists = new List<ReportLookup>();
            
        }

        /// <summary>
        /// Toggle report definition area
        /// </summary>
        /// <param name="value">enable state</param>
        public void EnableReportDefinitionArea(bool value)
        {
            this.comboBoxCategory.SelectedIndex = this.comboBoxCategoryDataArea.SelectedIndex = 0;
            this.comboBoxCategory.Enabled = this.comboBoxCategoryDataArea.Enabled = !value;            
        }

        #endregion

        #endregion

        #region Events
        /// <summary>
        /// Onload event
        /// </summary>
        /// <param name="args">EventArgs</param>
        protected override void OnLoad(EventArgs args)
        {
            base.OnLoad(args);
            if (!this.DesignMode)
            {
                // Set layout controls
                this.setLayoutControls();
                // Bind report definition data area
                this.bindReportDefinitionDataArea();
                
            }
        }

        /// <summary>
        /// Event handler for the comboBox selected index change 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void comboBoxCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            Category category = this.comboBoxCategory.SelectedItem as Category;
            if (category != null)
            {
                this.bindReportColumnsData(category);
            }
        }

        /// <summary>
        /// Event handler for the comboBox selected index change 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void comboBoxCategoryDataArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataArea dataArea = this.comboBoxCategoryDataArea.SelectedItem as DataArea;
            if (dataArea != null)
            {
                this.bindCategory(dataArea);
            }
        }

        /// <summary>
        /// Event to paint the report data area
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">painteventargs</param>
        private void panelDataArea_Paint(object sender, PaintEventArgs e)
        {
            this.panelDataArea.SuspendLayout();            
            int verticalScrollValue = this.panelDataArea.VerticalScroll.Value;
            int totalCols = this.m_ReportDetails.DisplayColumns.Count;
            int unscrolledCols = totalCols - this.scrollBarDataArea.Value;
            int numberOfRows = this.m_ReportDetails.DisplayStudents.Count;
            int horizontalAxisPosition = 0;
            int headerHeight = this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Height;
            int dividerFrequency = this.m_ReportDetails.ReportLayout.DividerFrequency;
            int fixedColumns = this.m_ReportDetails.ReportLayout.FixedColumns;            

            // Get the maximum level of nesting for any column
            int nestingPeak = this.getNestingPeak();
            // Get the no of row lines inside a cell as the maximum of  all rows inside a column's field value rows                                              
            int rowHeight = this.m_ReportDetails.ReportLayout.Font_ReportList_DataText.Height;
            // Draw column headings
            for (int displayColumnIndex = 0; displayColumnIndex < unscrolledCols; displayColumnIndex++)
            {
                int columnIndex = (displayColumnIndex < fixedColumns) ? displayColumnIndex : displayColumnIndex + this.scrollBarDataArea.Value;

                // Get Report column object
                ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];

                if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                    continue;

                Rectangle rect = new Rectangle(horizontalAxisPosition, 0, column.Width, rowHeight); 
                this.fillRectangle(e, rect);
                e.Graphics.Clip = new System.Drawing.Region(rect);

                string text = column.DisplayTitle;

                int titleoffset = 0;
                int sortstatus = this.m_ReportDetails.IsUnderlyingSortColumn(column);
                if (sortstatus != 0)
                {
                    string arrow = sortstatus == 1 ? "á" : "â";
                    Font wingdings = new System.Drawing.Font("WingDings", this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText.Size);
                    e.Graphics.DrawString(arrow, wingdings, System.Drawing.Brushes.Black, horizontalAxisPosition, 0);
                    titleoffset = (int)e.Graphics.MeasureString(arrow, wingdings).Width;
                }

                ReportColumn selectedColumn = null;

                if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex >= 0 && this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
                    selectedColumn = this.m_ReportDetails.DisplayColumns[this.m_ReportDetails.ReportLayout.SelectedColumnIndex];

                Brush brush = column == selectedColumn ? this.m_ReportDetails.ReportLayout.Brush_ColumnSelected : this.m_ReportDetails.ReportLayout.Brush_Heading;
                e.Graphics.DrawString(text, this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, brush, horizontalAxisPosition + titleoffset, 0);
               
                
                
                e.Graphics.ResetClip();
                horizontalAxisPosition += column.Width;
                // Set maximum rowspan value
                if (column.RowSpanAttribute.Value > this.m_RowSpan)
                {
                    this.m_RowSpan = column.RowSpanAttribute.Value;
                }
            }

            // Draw the header area for the complete display panel even if the summation of the display columns width is less than the display panel width
            Rectangle blankHeaderArea = new Rectangle(horizontalAxisPosition, 0, this.panelDataArea.Width - horizontalAxisPosition, rowHeight);
            this.fillRectangle(e, blankHeaderArea);
            e.Graphics.Clip = new System.Drawing.Region(blankHeaderArea);
            e.Graphics.DrawString("", this.m_ReportDetails.ReportLayout.Font_ReportColumn_HeaderText, this.m_ReportDetails.ReportLayout.Brush_Heading, horizontalAxisPosition, 0);
            e.Graphics.ResetClip();

            if (this.m_ReportDetails.DisplayColumns.Count == 0)
                return;

            // Draw student data
            int cellHeight = (nestingPeak + 1) * rowHeight * this.m_RowSpan;
            for (int studentIndex = 0; studentIndex < numberOfRows; studentIndex++)
            {
                int verticalAxisPosition = studentIndex * cellHeight + headerHeight - verticalScrollValue;
                if (verticalAxisPosition >= headerHeight && verticalAxisPosition < this.Height)
                {
                    StudentReportData student = this.m_ReportDetails.DisplayStudents[studentIndex];
                    horizontalAxisPosition = 0;
                    for (int displayColumnindex = 0; displayColumnindex < unscrolledCols; displayColumnindex++)
                    {
                        int columnIndex = (displayColumnindex < fixedColumns) ? displayColumnindex : displayColumnindex + this.scrollBarDataArea.Value;
                        // Draw Data
                        ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];

                        if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                            continue;

                        Rectangle rect = new Rectangle(horizontalAxisPosition, verticalAxisPosition, column.Width, cellHeight);

                        e.Graphics.Clip = new System.Drawing.Region(rect);
                        ReportColumn selectedColumn = null;

                        if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex >= 0 && this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.DisplayColumns.Count)
                            selectedColumn = this.m_ReportDetails.DisplayColumns[this.m_ReportDetails.ReportLayout.SelectedColumnIndex];

                        Brush brush = ((column == selectedColumn ? this.m_ReportDetails.ReportLayout.Brush_ColumnSelected :
                                        (student.Selected ? this.m_ReportDetails.ReportLayout.Brush_Data : this.m_ReportDetails.ReportLayout.Brush_RowSelected)));

                        // Get Column Data
                        string parentFieldText = student.GetColumnData(column);
                        e.Graphics.DrawString(parentFieldText, this.m_ReportDetails.ReportLayout.Font_ReportList_DataText, brush, horizontalAxisPosition, verticalAxisPosition);

                        if (this.m_ReportDetails.ReportLayout.IsNestingApplied && column.NestedColumns.Count > 0)
                        {
                            int verticalPosition = verticalAxisPosition;
                            foreach (ReportColumn nestedColumn in column.NestedColumns)
                            {
                                string nestedColumnFieldText = student.GetColumnData(nestedColumn);
                                verticalPosition += (rowHeight * this.m_RowSpan);
                                e.Graphics.DrawString(nestedColumnFieldText, this.m_ReportDetails.ReportLayout.Font_ReportList_DataText, brush, horizontalAxisPosition, verticalPosition);
                            }
                        }

                        e.Graphics.ResetClip();
                        // Draw Dividing lines
                        if (verticalAxisPosition > cellHeight && dividerFrequency > 0)
                        {
                            if (studentIndex % dividerFrequency == 0 && studentIndex != 0)
                            {
                                e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThinLine, new Point(0, verticalAxisPosition), new Point(this.panelDataArea.Width, verticalAxisPosition));
                            }
                        }
                        horizontalAxisPosition += column.Width;
                    }
                }
            }
            int bottomOfReport = numberOfRows * cellHeight + headerHeight - verticalScrollValue;
            // Draw vertical lines
            horizontalAxisPosition = 0;
            int fixedColumnIndex = fixedColumns;
            for (int displayColumnindex = 0; displayColumnindex < unscrolledCols; displayColumnindex++)
            {
                int columnIndex = (displayColumnindex < fixedColumns) ? displayColumnindex : displayColumnindex + this.scrollBarDataArea.Value;
                ReportColumn column = this.m_ReportDetails.DisplayColumns[columnIndex];
                if (column.NestedUnder != null && this.m_ReportDetails.ReportLayout.IsNestingApplied)
                {
                    fixedColumnIndex++;
                    continue;
                }
                Pen pen = (displayColumnindex == fixedColumnIndex) ? this.m_ReportDetails.ReportLayout.Pen_ColumnDividerLine : this.m_ReportDetails.ReportLayout.Pen_ThinLine;
                e.Graphics.DrawLine(pen, new Point(horizontalAxisPosition, 0), new Point(horizontalAxisPosition, bottomOfReport));
                horizontalAxisPosition += column.Width;
            }

            Pen drawingPen = (unscrolledCols == fixedColumnIndex) ? this.m_ReportDetails.ReportLayout.Pen_ColumnDividerLine : this.m_ReportDetails.ReportLayout.Pen_ThinLine;
            // Draw the vertical line after last display column
            e.Graphics.DrawLine(drawingPen, new Point(horizontalAxisPosition, 0), new Point(horizontalAxisPosition, bottomOfReport));
            // Draw the horizontal line at the end of the report
            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThinLine, new Point(0, bottomOfReport), new Point(this.panelDataArea.Width, bottomOfReport));
            // Draw the column header border
            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThickLine, new Point(0, 0), new Point(this.panelDataArea.Width, 0));
            e.Graphics.DrawLine(this.m_ReportDetails.ReportLayout.Pen_ThickLine, new Point(0, headerHeight), new Point(this.panelDataArea.Width, headerHeight));
            // Set the vertical scroll bar of the panel by positioning the report bottom panel
            if (this.panelReportBottom.Top != bottomOfReport)
            {
                this.panelReportBottom.Top = bottomOfReport;
            }
            this.panelColumnEdge.Height = bottomOfReport;
            this.panelDataArea.ResumeLayout();
        }

        /// <summary>
        /// Method to set the extended window style as WS_EX_COMPOSITED.
        /// </summary>
        protected override CreateParams CreateParams
        {
            // WS_EX_COMPOSITED : Windows XP style . Paints all descendants of a window in bottom-to-top painting order using double-buffering
            get
            {
                CreateParams cp = base.CreateParams;
                // Set the style as WS_EX_COMPOSITED
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }

        /// <summary>
        /// Drag enter event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">DragEventArgs</param>
        private void panelDataArea_DragEnter(object sender, DragEventArgs e)
        {
            if (this.DragMode != DragModeEnum.None)
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        /// <summary>
        /// Drag leave event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">DragEventArgs</param>
        private void panelDataArea_DragLeave(object sender, EventArgs e)
        {
            this.panelColumnEdge.Hide();
        }

        /// <summary>
        /// Drag enter event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">DragEventArgs</param>
        private void panelDataArea_MouseDown(object sender, MouseEventArgs e)
        {
            Point reportLayoutPoint = new Point(e.X, e.Y);
            // get location
            ReportControlHitInfo location = this.getReportDataAreaLocation(reportLayoutPoint);           

            if (location.InHeaderArea)
            {
                // Start "Re-order columns" drag operation                
                this.m_ReportDetails.ReportLayout.SelectedColumnIndex = location.ColumnIndex;
              
                // Raise the event from the container hosting this control
                if (this.SetColumnOptions != null)
                {
                    this.SetColumnOptions();
                }
                this.m_ReportDetails.ReportLayout.ColumnUnderDrag = location.Column;
                this.DragMode = DragModeEnum.RelocateColumn;
                this.panelDataArea.DoDragDrop(this.m_ReportDetails.ReportLayout.SelectedColumnIndex, DragDropEffects.Move);
            }

            if (e.Button == MouseButtons.Left || e.Button == MouseButtons.Right)
            {
                // get the selected row
                this.m_StudentReportData = null;
                if (location.RowIndex >= 0 && location.RowIndex < this.m_ReportDetails.DisplayStudents.Count)
                {
                    this.m_StudentReportData = this.m_ReportDetails.DisplayStudents[location.RowIndex];
                }

                if (e.Button == MouseButtons.Left)
                {
                    // Set the selected row focus
                    if (this.m_StudentReportData != null && !location.InHeaderArea)
                    {
                        this.setRowFocus(location);
                    }

                    if (Control.ModifierKeys == Keys.Control)
                    {
                        if (!location.Column.MultivalueColumn && location.Column.RowSpan < 2)
                        {
                            this.m_ReportDetails.SetUnderlyingSortColumn(location.Column);
                            this.m_ReportDetails.SortStudentList(this.m_ReportDetails.DisplayStudents);
                        }
                        else
                        {
                            MessageBox.Show(this, "Sorting is not permitted on this column!", "Student Lists", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                    if (location.NearDividerAtEnd >= 0)
                    {
                        // Start "Resize Column" drag operation
                        this.m_ReportDetails.ReportLayout.DragColumnDividerPosition = location.DividerPosition;
                        this.m_ReportDetails.ReportLayout.ColumnUnderDrag = this.m_ReportDetails.DisplayColumns[location.NearDividerAtEnd];
                        this.DragMode = DragModeEnum.ChangeColumnWidth;
                        this.panelDataArea.DoDragDrop(this.m_ReportDetails.ReportLayout.ColumnUnderDrag, DragDropEffects.Move);
                    }
                }
                else if (e.Button == MouseButtons.Right)
                {
                    Point mouseDownLocation = new Point(e.X, e.Y);
                    if (location.Column != null && location.InHeaderArea)
                    {
                        this.setContextMenusEnableState();
                        this.m_ContextMenuColumnOptions.Show(this.panelDataArea, mouseDownLocation);
                    }
                    else if (!location.InHeaderArea && this.m_StudentReportData != null && location.RowIndex >= 0)
                    {
                        this.m_ContextMenuStudentNavigationOptions.Show(this.panelDataArea, mouseDownLocation);
                    }
                }
            }
        }

        /// <external/>
        /// <summary>
        /// Scroll event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">ScrollEventArgs</param>
        private void panelDataArea_Scroll(object sender, ScrollEventArgs e)
        {
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Drag enter drag drop for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">DragEventArgs</param>
        private void panelDataArea_DragDrop(object sender, DragEventArgs e)
        {
            Point reportLayoutPoint = new Point(e.X, e.Y);
            reportLayoutPoint = this.panelDataArea.PointToClient(reportLayoutPoint);
            ReportControlHitInfo location = this.getReportDataAreaLocation(reportLayoutPoint);
            switch (this.DragMode)
            {
                case DragModeEnum.AddNewColumn:
                    // Get the target column index inside the report list on which the column under drag is dropped                       
                    int displayColumnIndex = location.ColumnIndex < 0 ? this.m_ReportDetails.DisplayColumns.Count : location.ColumnIndex;

                    // Don't drop multivalue columns into fixed columns area
                    if (this.m_ReportDetails.ReportLayout.ColumnUnderDrag.MultivalueColumn && displayColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns)
                    {
                        Cache.StatusMessage(MULTIVALUE_MESSAGE, UserMessageEventEnum.Error);
                        return;
                    }
                    // Check if the target column is inside fixed column's region and if the maximum range for fixed columns is execeeded 
                    else if (this.isFixedColumnRangeExceeded(displayColumnIndex))
                    {
                        Cache.StatusMessage(FIXED_COLUMNS_MESSAGE + this.m_ReportDetails.ReportLayout.MaximumFixedColumns.ToString() + COLUMN_DROP_MESSAGE, UserMessageEventEnum.Error);
                        return;
                    }
                    else
                    {
                        // Add new columns
                        this.addNewColumns(location);
                    }
                    break;
                case DragModeEnum.RelocateColumn:
                    if (location.ColumnIndex != this.m_ReportDetails.ReportLayout.SelectedColumnIndex)
                    {
                        // Check if the target column is inside fixed column's  region and if the maximum range for fixed columns is execeeded 
                        if (this.isFixedColumnRangeExceeded(location.ColumnIndex))
                        {
                            Cache.StatusMessage(FIXED_COLUMNS_MESSAGE + this.m_ReportDetails.ReportLayout.MaximumFixedColumns.ToString() + COLUMN_DROP_MESSAGE, UserMessageEventEnum.Error);
                            // Reset column selection
                            this.resetColumnSelection();
                            return;
                        }
                        this.m_ReportDetails.MoveColumn(this.m_ReportDetails.ReportLayout.SelectedColumnIndex, location.ColumnIndex);
                        int fixedColumnsCount = 0;
                        if (location.ColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns && location.ColumnIndex >= 0)
                        {
                            fixedColumnsCount++;
                        }
                        if (this.m_ReportDetails.ReportLayout.SelectedColumnIndex < this.m_ReportDetails.ReportLayout.FixedColumns)
                        {
                            fixedColumnsCount--;
                        }
                        this.m_ReportDetails.ReportLayout.FixedColumns += fixedColumnsCount;
                        // Display the fixed column value
                        if (this.DisplayFixedColumns != null)
                        {
                            this.DisplayFixedColumns();
                        }
                        // Reset column selection
                        this.resetColumnSelection();
                    }
                    break;
                case DragModeEnum.ChangeColumnWidth:
                    int newWidth = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Width + reportLayoutPoint.X - this.m_ReportDetails.ReportLayout.DragColumnDividerPosition;
                    if (newWidth < this.m_ReportDetails.ReportLayout.MinimumColumnWidth)
                    {
                        newWidth = this.m_ReportDetails.ReportLayout.MinimumColumnWidth;
                    }
                    this.m_ReportDetails.ReportLayout.ColumnUnderDrag.WidthAttribute.Value = newWidth;
                    this.panelColumnEdge.Hide();
                    this.resetColumnSelection();
                    break;
            }
            // Reset the scroll bounds
            this.SetHorizontalScrollBar();
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        ///  Mouse move  event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">MouseEventArgs</param>
        private void panelDataArea_MouseMove(object sender, MouseEventArgs e)
        {
            Point reportLayoutPoint = new Point(e.X, e.Y);
            ReportControlHitInfo location = this.getReportDataAreaLocation(reportLayoutPoint);
            if (location.NearDividerAtEnd >= 0)
            {
                this.panelDataArea.Cursor = Cursors.VSplit;
            }
            else
                if (location.InHeaderArea && location.Column != null)
                {
                    this.panelDataArea.Cursor = Cursors.Hand;
                }
                else
                {
                    this.panelDataArea.Cursor = Cursors.Default;
                }
        }

        /// <summary>
        /// Dragover event for the data area panel 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">DragEventArgs</param>
        private void panelDataArea_DragOver(object sender, DragEventArgs e)
        {
            Point reportLayoutPoint = new Point(e.X, e.Y);
            reportLayoutPoint = this.panelDataArea.PointToClient(reportLayoutPoint);
            ReportControlHitInfo location = this.getReportDataAreaLocation(reportLayoutPoint);
            if (this.DragMode == DragModeEnum.ChangeColumnWidth)
            {
                int newWidth = this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Width + reportLayoutPoint.X - this.m_ReportDetails.ReportLayout.DragColumnDividerPosition;
                if (newWidth < this.m_ReportDetails.ReportLayout.MinimumColumnWidth) newWidth = this.m_ReportDetails.ReportLayout.MinimumColumnWidth;
                this.panelColumnEdge.Left = this.m_ReportDetails.ReportLayout.DragColumnDividerPosition - this.m_ReportDetails.ReportLayout.ColumnUnderDrag.Width + newWidth;
                this.panelColumnEdge.Show();
            }
        }

        /// <summary>
        /// Context menu event handler
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemStudentDetails_Click(object sender, EventArgs e)
        {
            if (this.m_StudentReportData == null)
                return;

            StudentClassesUI studClasses = new StudentClassesUI(this.m_StudentReportData, SIMS.Entities.ReportLookupCache.EffectiveDate);
            AbstractComponentWindow abstractCompWindow = new AbstractComponentWindow(studClasses);
            abstractCompWindow.Show();
        }

        /// <summary>
        /// Context menu event handler
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemStudentTimeTable_Click(object sender, EventArgs e)
        {
            if (this.m_StudentReportData == null)
                return;

            TimetableUI studTimeTable = new TimetableUI(this.m_StudentReportData, SIMS.Entities.ReportLookupCache.EffectiveDate);
            AbstractComponentWindow abstractCompWindow = new AbstractComponentWindow(studTimeTable);
            abstractCompWindow.Show();
        }

        /// <summary>
        /// Context menu event handler to rename a column
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemRenameColumn_Click(object sender, EventArgs e)
        {
            this.RenameColumn();
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Context menu event handler to remove a column
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemRemoveColumn_Click(object sender, EventArgs e)
        {
            this.RemoveColumn();
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Context menu event handler to resize the column width
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemAutoSizeColumn_Click(object sender, EventArgs e)
        {
            this.AutoSizeColumn(this.m_ReportDetails.ReportLayout.SelectedColumnIndex);
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Context menu event handler to nest the column 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemNestColumn_Click(object sender, EventArgs e)
        {
            this.NestColumn(this.m_ReportDetails.ReportLayout.SelectedColumnIndex);
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Context menu event handler to remove column nesting
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">eventargs</param>
        private void menuItemRemoveColumnNesting_Click(object sender, EventArgs e)
        {
            this.RemoveColumnNesting(this.m_ReportDetails.ReportLayout.SelectedColumnIndex);
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Event handler for the horizontal scroll bar 
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">ScrollEventArgs</param>
        private void hScrollBarDataArea_Scroll(object sender, ScrollEventArgs e)
        {
            this.panelDataArea.Invalidate();
        }

        /// <summary>
        /// Event handler for the report column selection
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">MouseEventArgs</param>
        private void listViewReportColumns_MouseDown(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo listViewHitTestInfo = listViewReportColumns.HitTest(e.Location);
            if (listViewHitTestInfo.Item != null)
            {
                ReportColumn column = listViewHitTestInfo.Item.Tag as ReportColumn;
                if (column != null)
                {
                    this.m_ReportDetails.ReportLayout.ColumnUnderDrag = column;
                    this.DragMode = DragModeEnum.AddNewColumn;
                    this.listViewReportColumns.DoDragDrop(column, DragDropEffects.Move);
                }
            }
        }

        /// <summary>
        /// panelDataArea click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void panelDataArea_Click(object sender, EventArgs e)
        {
            this.panelDataArea.Focus();
        }

        /// <summary>
        /// Mousewheel event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void panelDataArea_MouseWheel(object sender, MouseEventArgs e)
        {
            if (this.panelDataArea.VerticalScroll.Visible)
                this.panelDataArea.VerticalScroll.Value = this.panelDataArea.VerticalScroll.Value + 10;
        }

        /// <summary>
        /// Event Handler for trapping key board events
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // we needs to set verticalscoll value 2 times to get latest scroll value
            if (!this.panelDataArea.VerticalScroll.Visible)
                return base.ProcessCmdKey(ref msg, keyData);

            if (keyData == Keys.Home)
            {
                this.panelDataArea.VerticalScroll.Value = this.panelDataArea.VerticalScroll.Minimum;
                this.panelDataArea.VerticalScroll.Value = this.panelDataArea.VerticalScroll.Minimum;
                this.panelDataArea.Invalidate();
                return true;
            }
            else if (keyData == Keys.End)
            {
                this.panelDataArea.VerticalScroll.Value = this.panelDataArea.VerticalScroll.Maximum;
                this.panelDataArea.VerticalScroll.Value = this.panelDataArea.VerticalScroll.Maximum;
                this.panelDataArea.Invalidate();
                return true;
            }
            else if (keyData == Keys.PageUp)
            {
                this.panelDataArea.VerticalScroll.Value = Math.Max(this.panelDataArea.VerticalScroll.Value - 250, this.panelDataArea.VerticalScroll.Minimum);
                this.panelDataArea.VerticalScroll.Value = Math.Max(this.panelDataArea.VerticalScroll.Value - 250, this.panelDataArea.VerticalScroll.Minimum);
                this.panelDataArea.Invalidate();
                return true;
            }
            else if (keyData == Keys.PageDown)
            {
                this.panelDataArea.VerticalScroll.Value = Math.Min(this.panelDataArea.VerticalScroll.Value + 250, this.panelDataArea.VerticalScroll.Maximum);
                this.panelDataArea.VerticalScroll.Value = Math.Min(this.panelDataArea.VerticalScroll.Value + 250, this.panelDataArea.VerticalScroll.Maximum);
                this.panelDataArea.Invalidate();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
    }
}