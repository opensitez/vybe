use crate::properties::PropertyBag;
use serde::{Deserialize, Serialize};
use uuid::Uuid;

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Bounds {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

impl Bounds {
    pub fn new(x: i32, y: i32, width: i32, height: i32) -> Self {
        Self { x, y, width, height }
    }

    pub fn contains(&self, x: i32, y: i32) -> bool {
        x >= self.x && x < self.x + self.width &&
        y >= self.y && y < self.y + self.height
    }

    pub fn right(&self) -> i32 {
        self.x + self.width
    }

    pub fn bottom(&self) -> i32 {
        self.y + self.height
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ControlType {
    Button,
    Label,
    TextBox,
    CheckBox,
    RadioButton,
    ComboBox,
    ListBox,
    Frame,
    PictureBox,
    RichTextBox,
    WebBrowser,
    TreeView,
    DataGridView,
    Panel,
    ListView,
    BindingNavigator,
    TabControl,
    TabPage,
    ProgressBar,
    NumericUpDown,
    MenuStrip,
    ToolStripMenuItem,
    ContextMenuStrip,
    StatusStrip,
    ToolStripStatusLabel,
    DateTimePicker,
    LinkLabel,
    ToolStrip,
    TrackBar,
    MaskedTextBox,
    SplitContainer,
    FlowLayoutPanel,
    TableLayoutPanel,
    MonthCalendar,
    HScrollBar,
    VScrollBar,
    ToolTip,
    // Non-visual data components (appear in component tray)
    BindingSourceComponent,
    DataSetComponent,
    DataTableComponent,
    DataAdapterComponent,
    // Arbitrary custom control type (fully qualified name)
    Custom(String),
}

impl ControlType {
    /// Parse a control type name (case-insensitive) into a ControlType variant.
    pub fn from_name(name: &str) -> Option<ControlType> {
        match name.to_lowercase().as_str() {
            "button" => Some(ControlType::Button),
            "label" => Some(ControlType::Label),
            "textbox" => Some(ControlType::TextBox),
            "checkbox" => Some(ControlType::CheckBox),
            "radiobutton" => Some(ControlType::RadioButton),
            "combobox" => Some(ControlType::ComboBox),
            "listbox" => Some(ControlType::ListBox),
            "frame" | "groupbox" => Some(ControlType::Frame),
            "picturebox" => Some(ControlType::PictureBox),
            "richtextbox" => Some(ControlType::RichTextBox),
            "webbrowser" => Some(ControlType::WebBrowser),
            "treeview" => Some(ControlType::TreeView),
            "datagridview" => Some(ControlType::DataGridView),
            "panel" => Some(ControlType::Panel),
            "listview" => Some(ControlType::ListView),
            "bindingnavigator" => Some(ControlType::BindingNavigator),
            "tabcontrol" => Some(ControlType::TabControl),
            "tabpage" => Some(ControlType::TabPage),
            "progressbar" => Some(ControlType::ProgressBar),
            "numericupdown" => Some(ControlType::NumericUpDown),
            "menustrip" => Some(ControlType::MenuStrip),
            "toolstripmenuitem" => Some(ControlType::ToolStripMenuItem),
            "contextmenustrip" => Some(ControlType::ContextMenuStrip),
            "statusstrip" => Some(ControlType::StatusStrip),
            "toolstripstatuslabel" => Some(ControlType::ToolStripStatusLabel),
            "datetimepicker" => Some(ControlType::DateTimePicker),
            "linklabel" => Some(ControlType::LinkLabel),
            "toolstrip" => Some(ControlType::ToolStrip),
            "trackbar" => Some(ControlType::TrackBar),
            "maskedtextbox" => Some(ControlType::MaskedTextBox),
            "splitcontainer" => Some(ControlType::SplitContainer),
            "flowlayoutpanel" => Some(ControlType::FlowLayoutPanel),
            "tablelayoutpanel" => Some(ControlType::TableLayoutPanel),
            "monthcalendar" => Some(ControlType::MonthCalendar),
            "hscrollbar" => Some(ControlType::HScrollBar),
            "vscrollbar" => Some(ControlType::VScrollBar),
            "tooltip" => Some(ControlType::ToolTip),
            _ => Some(ControlType::Custom(name.to_string())),
        }
    }

    pub fn as_str(&self) -> &str {
        match self {
            ControlType::Button => "Button",
            ControlType::Label => "Label",
            ControlType::TextBox => "TextBox",
            ControlType::CheckBox => "CheckBox",
            ControlType::RadioButton => "RadioButton",
            ControlType::ComboBox => "ComboBox",
            ControlType::ListBox => "ListBox",
            ControlType::Frame => "Frame",
            ControlType::PictureBox => "PictureBox",
            ControlType::RichTextBox => "RichTextBox",
            ControlType::WebBrowser => "WebBrowser",
            ControlType::TreeView => "TreeView",
            ControlType::DataGridView => "DataGridView",
            ControlType::Panel => "Panel",
            ControlType::ListView => "ListView",
            ControlType::BindingNavigator => "BindingNavigator",
            ControlType::TabControl => "TabControl",
            ControlType::TabPage => "TabPage",
            ControlType::ProgressBar => "ProgressBar",
            ControlType::NumericUpDown => "NumericUpDown",
            ControlType::MenuStrip => "MenuStrip",
            ControlType::ToolStripMenuItem => "ToolStripMenuItem",
            ControlType::ContextMenuStrip => "ContextMenuStrip",
            ControlType::StatusStrip => "StatusStrip",
            ControlType::ToolStripStatusLabel => "ToolStripStatusLabel",
            ControlType::DateTimePicker => "DateTimePicker",
            ControlType::LinkLabel => "LinkLabel",
            ControlType::ToolStrip => "ToolStrip",
            ControlType::TrackBar => "TrackBar",
            ControlType::MaskedTextBox => "MaskedTextBox",
            ControlType::SplitContainer => "SplitContainer",
            ControlType::FlowLayoutPanel => "FlowLayoutPanel",
            ControlType::TableLayoutPanel => "TableLayoutPanel",
            ControlType::MonthCalendar => "MonthCalendar",
            ControlType::HScrollBar => "HScrollBar",
            ControlType::VScrollBar => "VScrollBar",
            ControlType::ToolTip => "ToolTip",
            ControlType::BindingSourceComponent => "BindingSource",
            ControlType::DataSetComponent => "DataSet",
            ControlType::DataTableComponent => "DataTable",
            ControlType::DataAdapterComponent => "DataAdapter",
            ControlType::Custom(s) => s.as_str(),
        }
    }

    /// Returns true if this is a non-visual component (lives in component tray, not form surface)
    pub fn is_non_visual(&self) -> bool {
        matches!(self,
            ControlType::BindingSourceComponent |
            ControlType::DataSetComponent |
            ControlType::DataTableComponent |
            ControlType::DataAdapterComponent
        )
    }

    /// Returns true if this control type supports DataSource/DataMember complex binding
    /// (list/grid controls that display multiple records from a BindingSource)
    pub fn supports_complex_binding(&self) -> bool {
        matches!(self,
            ControlType::DataGridView |
            ControlType::ListBox |
            ControlType::ComboBox |
            ControlType::BindingNavigator
        )
    }

    pub fn default_name_prefix(&self) -> &str {
        match self {
            ControlType::Button => "btn",
            ControlType::Label => "lbl",
            ControlType::TextBox => "txt",
            ControlType::CheckBox => "chk",
            ControlType::RadioButton => "opt",
            ControlType::ComboBox => "cbo",
            ControlType::ListBox => "lst",
            ControlType::Frame => "fra",
            ControlType::PictureBox => "pic",
            ControlType::RichTextBox => "rtf",
            ControlType::WebBrowser => "web",
            ControlType::TreeView => "tvw",
            ControlType::DataGridView => "dgv",
            ControlType::Panel => "pnl",
            ControlType::ListView => "lvw",
            ControlType::BindingNavigator => "bnav",
            ControlType::TabControl => "tab",
            ControlType::TabPage => "tp",
            ControlType::ProgressBar => "pb",
            ControlType::NumericUpDown => "nud",
            ControlType::MenuStrip => "ms",
            ControlType::ToolStripMenuItem => "tsmi",
            ControlType::ContextMenuStrip => "cms",
            ControlType::StatusStrip => "ss",
            ControlType::ToolStripStatusLabel => "tssl",
            ControlType::DateTimePicker => "dtp",
            ControlType::LinkLabel => "lnk",
            ControlType::ToolStrip => "ts",
            ControlType::TrackBar => "trk",
            ControlType::MaskedTextBox => "mtxt",
            ControlType::SplitContainer => "sc",
            ControlType::FlowLayoutPanel => "flp",
            ControlType::TableLayoutPanel => "tlp",
            ControlType::MonthCalendar => "mc",
            ControlType::HScrollBar => "hsb",
            ControlType::VScrollBar => "vsb",
            ControlType::ToolTip => "tt",
            ControlType::BindingSourceComponent => "bs",
            ControlType::DataSetComponent => "ds",
            ControlType::DataTableComponent => "dt",
            ControlType::DataAdapterComponent => "da",
            ControlType::Custom(_) => "ctrl",
        }
    }

    pub fn default_size(&self) -> (i32, i32) {
        match self {
            ControlType::Button => (120, 30),
            ControlType::Label => (80, 20),
            ControlType::TextBox => (150, 25),
            ControlType::CheckBox => (120, 20),
            ControlType::RadioButton => (120, 20),
            ControlType::ComboBox => (150, 25),
            ControlType::ListBox => (150, 100),
            ControlType::Frame => (200, 150),
            ControlType::PictureBox => (100, 100),
            ControlType::RichTextBox => (200, 150),
            ControlType::WebBrowser => (400, 300),
            ControlType::TreeView => (200, 200),
            ControlType::DataGridView => (300, 200),
            ControlType::Panel => (200, 150),
            ControlType::ListView => (250, 200),
            ControlType::BindingNavigator => (300, 25),
            ControlType::TabControl => (300, 200),
            ControlType::TabPage => (300, 200),
            ControlType::ProgressBar => (200, 23),
            ControlType::NumericUpDown => (120, 23),
            ControlType::MenuStrip => (300, 24),
            ControlType::ToolStripMenuItem => (100, 22),
            ControlType::ContextMenuStrip => (150, 24),
            ControlType::StatusStrip => (300, 22),
            ControlType::ToolStripStatusLabel => (100, 22),
            ControlType::DateTimePicker => (200, 23),
            ControlType::LinkLabel => (100, 20),
            ControlType::ToolStrip => (300, 25),
            ControlType::TrackBar => (200, 45),
            ControlType::MaskedTextBox => (150, 23),
            ControlType::SplitContainer => (300, 200),
            ControlType::FlowLayoutPanel => (200, 150),
            ControlType::TableLayoutPanel => (200, 150),
            ControlType::MonthCalendar => (227, 164),
            ControlType::HScrollBar => (200, 17),
            ControlType::VScrollBar => (17, 200),
            ControlType::ToolTip => (32, 32),
            ControlType::BindingSourceComponent => (32, 32),
            ControlType::DataSetComponent => (32, 32),
            ControlType::DataTableComponent => (32, 32),
            ControlType::DataAdapterComponent => (32, 32),
            ControlType::Custom(_) => (100, 100),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Control {
    pub id: Uuid,
    pub name: String,
    pub control_type: ControlType,
    pub bounds: Bounds,
    pub properties: PropertyBag,
    pub tab_index: i32,
    #[serde(default)]
    pub parent_id: Option<Uuid>,
    #[serde(default)]
    pub index: Option<i32>,
}

impl Control {
    pub fn new(control_type: ControlType, name: String, x: i32, y: i32) -> Self {
        let (width, height) = control_type.default_size();
        let mut properties = PropertyBag::new();

        // Set default properties based on control type
        match control_type {
            ControlType::Button => {
                properties.set("Text", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::Label => {
                properties.set("Text", name.clone());
                properties.set("Visible", true);
            }
            ControlType::TextBox => {
                properties.set("Text", "");
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::CheckBox => {
                properties.set("Text", name.clone());
                properties.set("Checked", false);
                use crate::properties::PropertyValue;
                properties.set_raw("CheckState", PropertyValue::Integer(0));
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::RadioButton => {
                properties.set("Text", name.clone());
                properties.set("Checked", false);
                use crate::properties::PropertyValue;
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::ComboBox | ControlType::ListBox => {
                use crate::properties::PropertyValue;
                properties.set("Enabled", true);
                properties.set("Visible", true);
                // Initialize with empty list
                properties.set_raw("List", PropertyValue::StringArray(vec![]));
                properties.set_raw("ListValues", PropertyValue::StringArray(vec![]));
                    properties.set_raw("ListIndex", PropertyValue::Integer(-1));
                    properties.set_raw("Text", PropertyValue::String(String::new()));
                    properties.set_raw("Value", PropertyValue::String(String::new()));
                // ComboBox-specific: DropDownStyle default is DropDown (1) in .NET
                if control_type == ControlType::ComboBox {
                    properties.set_raw("DropDownStyle", PropertyValue::Integer(1));
                }
            }
            ControlType::RichTextBox => {
                properties.set("Text", "");
                properties.set("HTML", ""); // HTML content
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("ReadOnly", false);
                properties.set("ScrollBars", 2); // 2 = Both
                properties.set("ToolbarVisible", true); // Show toolbar by default
            }
            ControlType::WebBrowser => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("URL", "about:blank");
            }
            ControlType::ListView => {
                properties.set("View", "Details"); // Report, List, Icon, SmallIcon
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::TreeView => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("PathSeparator", "\\");
            }
            ControlType::DataGridView => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("AllowUserToAddRows", true);
                properties.set("AllowUserToDeleteRows", true);
            }
            ControlType::Panel => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("BorderStyle", "None"); // None, FixedSingle, Fixed3D
            }
            ControlType::BindingNavigator => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("BindingSource", "");
            }
            ControlType::TabControl => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("SelectedIndex", 0);
            }
            ControlType::TabPage => {
                properties.set("Text", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::ProgressBar => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                use crate::properties::PropertyValue;
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set_raw("Minimum", PropertyValue::Integer(0));
                properties.set_raw("Maximum", PropertyValue::Integer(100));
                properties.set_raw("Step", PropertyValue::Integer(10));
            }
            ControlType::NumericUpDown => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                use crate::properties::PropertyValue;
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set_raw("Minimum", PropertyValue::Integer(0));
                properties.set_raw("Maximum", PropertyValue::Integer(100));
                properties.set_raw("Increment", PropertyValue::Integer(1));
                properties.set_raw("DecimalPlaces", PropertyValue::Integer(0));
            }
            ControlType::MenuStrip => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::ToolStripMenuItem => {
                properties.set("Text", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::ContextMenuStrip => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::StatusStrip => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::ToolStripStatusLabel => {
                properties.set("Text", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::DateTimePicker => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("Format", "Long");
                properties.set("CustomFormat", "");
                properties.set("ShowCheckBox", false);
                properties.set("Checked", true);
            }
            ControlType::LinkLabel => {
                properties.set("Text", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("LinkColor", "#0066cc");
                properties.set("VisitedLinkColor", "#800080");
                properties.set("LinkVisited", false);
            }
            ControlType::ToolStrip => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::TrackBar => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                use crate::properties::PropertyValue;
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set_raw("Minimum", PropertyValue::Integer(0));
                properties.set_raw("Maximum", PropertyValue::Integer(10));
                properties.set_raw("TickFrequency", PropertyValue::Integer(1));
                properties.set_raw("SmallChange", PropertyValue::Integer(1));
                properties.set_raw("LargeChange", PropertyValue::Integer(5));
                properties.set("Orientation", "Horizontal");
            }
            ControlType::MaskedTextBox => {
                properties.set("Text", "");
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("Mask", "");
                properties.set("PromptChar", "_");
            }
            ControlType::SplitContainer => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("Orientation", "Vertical");
                use crate::properties::PropertyValue;
                properties.set_raw("SplitterDistance", PropertyValue::Integer(100));
            }
            ControlType::FlowLayoutPanel => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("FlowDirection", "LeftToRight");
                properties.set("WrapContents", true);
                properties.set("BorderStyle", "None");
            }
            ControlType::TableLayoutPanel => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                use crate::properties::PropertyValue;
                properties.set_raw("ColumnCount", PropertyValue::Integer(2));
                properties.set_raw("RowCount", PropertyValue::Integer(2));
                properties.set("BorderStyle", "None");
            }
            ControlType::MonthCalendar => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                properties.set("ShowToday", true);
                properties.set("ShowTodayCircle", true);
                properties.set("ShowWeekNumbers", false);
            }
            ControlType::HScrollBar | ControlType::VScrollBar => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
                use crate::properties::PropertyValue;
                properties.set_raw("Value", PropertyValue::Integer(0));
                properties.set_raw("Minimum", PropertyValue::Integer(0));
                properties.set_raw("Maximum", PropertyValue::Integer(100));
                properties.set_raw("SmallChange", PropertyValue::Integer(1));
                properties.set_raw("LargeChange", PropertyValue::Integer(10));
            }
            ControlType::ToolTip => {
                properties.set("Active", true);
                use crate::properties::PropertyValue;
                properties.set_raw("AutoPopDelay", PropertyValue::Integer(5000));
                properties.set_raw("InitialDelay", PropertyValue::Integer(500));
                properties.set_raw("ReshowDelay", PropertyValue::Integer(100));
                properties.set("ShowAlways", false);
            }
            ControlType::BindingSourceComponent => {
                properties.set("DataSource", "");
                properties.set("DataMember", "");
                properties.set("Filter", "");
                properties.set("Sort", "");
            }
            ControlType::DataSetComponent => {
                properties.set("DataSetName", "NewDataSet");
            }
            ControlType::DataTableComponent => {
                properties.set("TableName", "Table1");
            }
            ControlType::DataAdapterComponent => {
                properties.set("SelectCommand", "");
                properties.set("ConnectionString", "");
            }
            ControlType::Custom(_) => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            _ => {
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
        }

        // Common appearance defaults
        properties.set("BackColor", "#f8fafc");
        properties.set("ForeColor", "#0f172a");
        properties.set("Font", "Segoe UI, 12px");

        Self {
            id: Uuid::new_v4(),
            name,
            control_type,
            bounds: Bounds::new(x, y, width, height),
            properties,
            tab_index: 0,
            parent_id: None,
            index: None,
        }
    }

    pub fn get_text(&self) -> Option<&str> {
        self.properties.get_string("Text")
    }

    pub fn set_text(&mut self, text: impl Into<String>) {
        self.properties.set("Text", text.into());
    }

    pub fn is_enabled(&self) -> bool {
        self.properties.get_bool("Enabled").unwrap_or(true)
    }

    pub fn set_enabled(&mut self, enabled: bool) {
        self.properties.set("Enabled", enabled);
    }

    pub fn is_visible(&self) -> bool {
        self.properties.get_bool("Visible").unwrap_or(true)
    }

    pub fn set_visible(&mut self, visible: bool) {
        self.properties.set("Visible", visible);
    }

    pub fn get_back_color(&self) -> Option<&str> {
        self.properties.get_string("BackColor")
    }

    pub fn set_back_color(&mut self, color: impl Into<String>) {
        self.properties.set("BackColor", color.into());
    }

    pub fn get_fore_color(&self) -> Option<&str> {
        self.properties.get_string("ForeColor")
    }

    pub fn set_fore_color(&mut self, color: impl Into<String>) {
        self.properties.set("ForeColor", color.into());
    }

    pub fn get_font(&self) -> Option<&str> {
        self.properties.get_string("Font")
    }

    pub fn set_font(&mut self, font: impl Into<String>) {
        self.properties.set("Font", font.into());
    }
    
    pub fn get_list_items(&self) -> Vec<String> {
        self.properties.get_string_array("List")
            .map(|arr| arr.clone())
            .unwrap_or_default()
    }
    
    pub fn set_list_items(&mut self, items: Vec<String>) {
        use crate::properties::PropertyValue;
        self.properties.set_raw("List", PropertyValue::StringArray(items));
    }

    pub fn display_name(&self) -> String {
        if let Some(idx) = self.index {
            format!("{}({})", self.name, idx)
        } else {
            self.name.clone()
        }
    }

    pub fn is_array_member(&self) -> bool {
        self.index.is_some()
    }
}
