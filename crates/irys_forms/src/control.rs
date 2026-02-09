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

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
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
}

impl ControlType {
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
        }
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
                properties.set("Caption", name.clone());
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::Label => {
                properties.set("Caption", name.clone());
                properties.set("Visible", true);
            }
            ControlType::TextBox => {
                properties.set("Text", "");
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::CheckBox => {
                properties.set("Caption", name.clone());
                properties.set("Value", 0);
                properties.set("Enabled", true);
                properties.set("Visible", true);
            }
            ControlType::RadioButton => {
                properties.set("Caption", name.clone());
                properties.set("Value", 0);
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

    pub fn get_caption(&self) -> Option<&str> {
        self.properties.get_string("Caption")
    }

    pub fn set_caption(&mut self, caption: impl Into<String>) {
        self.properties.set("Caption", caption.into());
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
