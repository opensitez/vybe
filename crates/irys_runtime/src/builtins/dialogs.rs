use crate::{RuntimeError, Value};
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

/// Represents a File Dialog (OpenFileDialog or SaveFileDialog)
#[derive(Debug, Clone)]
pub struct FileDialogData {
    pub filter: String,
    pub filter_index: i32,
    pub initial_directory: String,
    pub file_name: String,
    pub title: String,
    pub multiselect: bool,
    pub selected_files: Vec<String>,
    pub dialog_result: bool,
}

impl Default for FileDialogData {
    fn default() -> Self {
        Self {
            filter: String::new(),
            filter_index: 1,
            initial_directory: String::new(),
            file_name: String::new(),
            title: String::new(),
            multiselect: false,
            selected_files: Vec::new(),
            dialog_result: false,
        }
    }
}

/// Represents a Color Dialog
#[derive(Debug, Clone)]
pub struct ColorDialogData {
    pub color: String,           // Hex color like "#FF0000"
    pub full_open: bool,         // Whether to show full color picker
    pub any_color: bool,         // Allow any color
    pub solid_color_only: bool,  // Only solid colors
    pub custom_colors: Vec<String>,
    pub dialog_result: bool,
}

impl Default for ColorDialogData {
    fn default() -> Self {
        Self {
            color: "#000000".to_string(),
            full_open: false,
            any_color: false,
            solid_color_only: false,
            custom_colors: Vec::new(),
            dialog_result: false,
        }
    }
}

/// Represents a Font Dialog
#[derive(Debug, Clone)]
pub struct FontDialogData {
    pub font_name: String,
    pub font_size: f32,
    pub font_bold: bool,
    pub font_italic: bool,
    pub font_underline: bool,
    pub font_strikeout: bool,
    pub color: String,
    pub dialog_result: bool,
}

impl Default for FontDialogData {
    fn default() -> Self {
        Self {
            font_name: "Arial".to_string(),
            font_size: 12.0,
            font_bold: false,
            font_italic: false,
            font_underline: false,
            font_strikeout: false,
            color: "#000000".to_string(),
            dialog_result: false,
        }
    }
}

/// Represents a Folder Browser Dialog
#[derive(Debug, Clone)]
pub struct FolderBrowserDialogData {
    pub selected_path: String,
    pub description: String,
    pub root_folder: String,
    pub show_new_folder_button: bool,
    pub dialog_result: bool,
}

impl Default for FolderBrowserDialogData {
    fn default() -> Self {
        Self {
            selected_path: String::new(),
            description: String::new(),
            root_folder: String::new(),
            show_new_folder_button: true,
            dialog_result: false,
        }
    }
}

/// Creates an OpenFileDialog object
pub fn create_openfiledialog() -> Value {
    let data = FileDialogData::default();
    let mut fields = HashMap::new();
    fields.insert("Filter".to_string(), Value::String(data.filter.clone()));
    fields.insert("FilterIndex".to_string(), Value::Integer(data.filter_index));
    fields.insert("InitialDirectory".to_string(), Value::String(data.initial_directory.clone()));
    fields.insert("FileName".to_string(), Value::String(data.file_name.clone()));
    fields.insert("Title".to_string(), Value::String(data.title.clone()));
    fields.insert("Multiselect".to_string(), Value::Boolean(data.multiselect));
    fields.insert("_dialog_type".to_string(), Value::String("OpenFileDialog".to_string()));
    
    Value::Object(Rc::new(RefCell::new(crate::ObjectData {
        class_name: "OpenFileDialog".to_string(),
        fields,
    })))
}

/// Creates a SaveFileDialog object
pub fn create_savefiledialog() -> Value {
    let data = FileDialogData::default();
    let mut fields = HashMap::new();
    fields.insert("Filter".to_string(), Value::String(data.filter.clone()));
    fields.insert("FilterIndex".to_string(), Value::Integer(data.filter_index));
    fields.insert("InitialDirectory".to_string(), Value::String(data.initial_directory.clone()));
    fields.insert("FileName".to_string(), Value::String(data.file_name.clone()));
    fields.insert("Title".to_string(), Value::String(data.title.clone()));
    fields.insert("_dialog_type".to_string(), Value::String("SaveFileDialog".to_string()));
    
    Value::Object(Rc::new(RefCell::new(crate::ObjectData {
        class_name: "SaveFileDialog".to_string(),
        fields,
    })))
}

/// Creates a ColorDialog object
pub fn create_colordialog() -> Value {
    let data = ColorDialogData::default();
    let mut fields = HashMap::new();
    fields.insert("Color".to_string(), Value::String(data.color.clone()));
    fields.insert("FullOpen".to_string(), Value::Boolean(data.full_open));
    fields.insert("AnyColor".to_string(), Value::Boolean(data.any_color));
    fields.insert("SolidColorOnly".to_string(), Value::Boolean(data.solid_color_only));
    fields.insert("_dialog_type".to_string(), Value::String("ColorDialog".to_string()));
    
    Value::Object(Rc::new(RefCell::new(crate::ObjectData {
        class_name: "ColorDialog".to_string(),
        fields,
    })))
}

/// Creates a FontDialog object
pub fn create_fontdialog() -> Value {
    let data = FontDialogData::default();
    let mut fields = HashMap::new();
    fields.insert("Font".to_string(), Value::String(format!("{}, {}pt", data.font_name, data.font_size)));
    fields.insert("Color".to_string(), Value::String(data.color.clone()));
    fields.insert("_dialog_type".to_string(), Value::String("FontDialog".to_string()));
    
    Value::Object(Rc::new(RefCell::new(crate::ObjectData {
        class_name: "FontDialog".to_string(),
        fields,
    })))
}

/// Creates a FolderBrowserDialog object
pub fn create_folderbrowserdialog() -> Value {
    let data = FolderBrowserDialogData::default();
    let mut fields = HashMap::new();
    fields.insert("SelectedPath".to_string(), Value::String(data.selected_path.clone()));
    fields.insert("Description".to_string(), Value::String(data.description.clone()));
    fields.insert("ShowNewFolderButton".to_string(), Value::Boolean(data.show_new_folder_button));
    fields.insert("_dialog_type".to_string(), Value::String("FolderBrowserDialog".to_string()));
    
    Value::Object(Rc::new(RefCell::new(crate::ObjectData {
        class_name: "FolderBrowserDialog".to_string(),
        fields,
    })))
}

/// ShowDialog method for file dialogs - Opens REAL native file/folder dialogs
pub fn dialog_showdialog(dialog: &Value) -> Result<Value, RuntimeError> {
    match dialog {
        Value::Object(obj_ref) => {
            let dialog_type = obj_ref.borrow().fields.get("_dialog_type")
                .map(|v| v.as_string().to_string())
                .unwrap_or_default();
            
            match dialog_type.as_str() {
                "OpenFileDialog" => {
                    let filter = obj_ref.borrow().fields.get("Filter")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    let title = obj_ref.borrow().fields.get("Title")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_else(|| "Open File".to_string());
                    let initial_dir = obj_ref.borrow().fields.get("InitialDirectory")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    
                    let mut dialog = rfd::FileDialog::new().set_title(&title);
                    
                    if !initial_dir.is_empty() {
                        dialog = dialog.set_directory(&initial_dir);
                    }
                    
                    // Parse filter: "Text Files|*.txt|All Files|*.*"
                    if !filter.is_empty() {
                        let parts: Vec<&str> = filter.split('|').collect();
                        for i in (0..parts.len()).step_by(2) {
                            if i + 1 < parts.len() {
                                let name = parts[i];
                                let pattern = parts[i + 1];
                                let exts: Vec<&str> = pattern.split(';')
                                    .map(|p| p.trim().trim_start_matches("*.").trim_start_matches('*'))
                                    .filter(|e| !e.is_empty() && *e != ".*")
                                    .collect();
                                if !exts.is_empty() {
                                    dialog = dialog.add_filter(name, &exts);
                                }
                            }
                        }
                    }
                    
                    if let Some(path) = dialog.pick_file() {
                        obj_ref.borrow_mut().fields.insert("FileName".to_string(), 
                            Value::String(path.to_string_lossy().to_string()));
                        return Ok(Value::Boolean(true));
                    }
                    Ok(Value::Boolean(false))
                }
                
                "SaveFileDialog" => {
                    let filter = obj_ref.borrow().fields.get("Filter")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    let title = obj_ref.borrow().fields.get("Title")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_else(|| "Save File".to_string());
                    let initial_dir = obj_ref.borrow().fields.get("InitialDirectory")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    let file_name = obj_ref.borrow().fields.get("FileName")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    
                    let mut dialog = rfd::FileDialog::new().set_title(&title);
                    
                    if !initial_dir.is_empty() {
                        dialog = dialog.set_directory(&initial_dir);
                    }
                    
                    if !file_name.is_empty() {
                        dialog = dialog.set_file_name(&file_name);
                    }
                    
                    // Parse filter
                    if !filter.is_empty() {
                        let parts: Vec<&str> = filter.split('|').collect();
                        for i in (0..parts.len()).step_by(2) {
                            if i + 1 < parts.len() {
                                let name = parts[i];
                                let pattern = parts[i + 1];
                                let exts: Vec<&str> = pattern.split(';')
                                    .map(|p| p.trim().trim_start_matches("*.").trim_start_matches('*'))
                                    .filter(|e| !e.is_empty() && *e != ".*")
                                    .collect();
                                if !exts.is_empty() {
                                    dialog = dialog.add_filter(name, &exts);
                                }
                            }
                        }
                    }
                    
                    if let Some(path) = dialog.save_file() {
                        obj_ref.borrow_mut().fields.insert("FileName".to_string(), 
                            Value::String(path.to_string_lossy().to_string()));
                        return Ok(Value::Boolean(true));
                    }
                    Ok(Value::Boolean(false))
                }
                
                "FolderBrowserDialog" => {
                    let description = obj_ref.borrow().fields.get("Description")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_else(|| "Select Folder".to_string());
                    let initial_dir = obj_ref.borrow().fields.get("SelectedPath")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_default();
                    
                    let mut dialog = rfd::FileDialog::new().set_title(&description);
                    
                    if !initial_dir.is_empty() {
                        dialog = dialog.set_directory(&initial_dir);
                    }
                    
                    if let Some(path) = dialog.pick_folder() {
                        obj_ref.borrow_mut().fields.insert("SelectedPath".to_string(), 
                            Value::String(path.to_string_lossy().to_string()));
                        return Ok(Value::Boolean(true));
                    }
                    Ok(Value::Boolean(false))
                }
                
                "ColorDialog" => {
                    let current_color = obj_ref.borrow().fields.get("Color")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_else(|| "#000000".to_string());
                    
                    // Use native OS color picker
                    if let Some(selected_color) = show_native_color_picker(&current_color) {
                        obj_ref.borrow_mut().fields.insert("Color".to_string(), 
                            Value::String(selected_color));
                        Ok(Value::Boolean(true))
                    } else {
                        Ok(Value::Boolean(false))
                    }
                }
                
                "FontDialog" => {
                    let current_font = obj_ref.borrow().fields.get("FontName")
                        .map(|v| v.as_string().to_string())
                        .unwrap_or_else(|| "Arial".to_string());
                    let current_size = obj_ref.borrow().fields.get("FontSize")
                        .and_then(|v| match v {
                            Value::Double(d) => Some(*d),
                            Value::Single(f) => Some(*f as f64),
                            Value::Integer(i) => Some(*i as f64),
                            _ => None,
                        })
                        .unwrap_or(12.0);
                    
                    // Use native OS font picker
                    if let Some((font_name, font_size)) = show_native_font_picker(&current_font, current_size) {
                        obj_ref.borrow_mut().fields.insert("FontName".to_string(), 
                            Value::String(font_name));
                        obj_ref.borrow_mut().fields.insert("FontSize".to_string(), 
                            Value::Double(font_size));
                        Ok(Value::Boolean(true))
                    } else {
                        Ok(Value::Boolean(false))
                    }
                }
                
                _ => Ok(Value::Boolean(false))
            }
        }
        _ => Err(RuntimeError::Custom("ShowDialog called on non-dialog object".to_string()))
    }
}

/// MsgBox function is already implemented in msgbox.rs
/// InputBox function is already implemented in info_fns.rs
/// DialogResult constants

pub const DIALOG_RESULT_OK: i32 = 1;
pub const DIALOG_RESULT_CANCEL: i32 = 2;
pub const DIALOG_RESULT_ABORT: i32 = 3;
pub const DIALOG_RESULT_RETRY: i32 = 4;
pub const DIALOG_RESULT_IGNORE: i32 = 5;
pub const DIALOG_RESULT_YES: i32 = 6;
pub const DIALOG_RESULT_NO: i32 = 7;

/// Show native OS color picker and return selected color as hex string
fn show_native_color_picker(default_color: &str) -> Option<String> {
    #[cfg(target_os = "macos")]
    {
        use std::process::Command;
        
        // Parse hex color to RGB 0-65535 for AppleScript
        let rgb = parse_hex_to_applescript_rgb(default_color);
        
        eprintln!("[ColorDialog] Default color: {} -> RGB({}, {}, {})", default_color, rgb.0, rgb.1, rgb.2);
        
        let script = format!(
            "choose color default color {{{}, {}, {}}}",
            rgb.0, rgb.1, rgb.2
        );
        
        match Command::new("osascript").arg("-e").arg(&script).output() {
            Ok(output) if output.status.success() => {
                let result = String::from_utf8_lossy(&output.stdout);
                eprintln!("[ColorDialog] AppleScript result: {:?}", result);
                // Parse result like "65535, 0, 0" to #FF0000
                let hex = applescript_rgb_to_hex(&result);
                eprintln!("[ColorDialog] Parsed hex: {:?}", hex);
                hex
            }
            Err(e) => {
                eprintln!("[ColorDialog] Error running osascript: {}", e);
                None
            }
            _ => {
                eprintln!("[ColorDialog] osascript returned non-success status");
                None
            }
        }
    }
    
    #[cfg(target_os = "linux")]
    {
        use std::process::Command;
        
        match Command::new("zenity")
            .arg("--color-selection")
            .arg(format!("--color={}", default_color))
            .output()
        {
            Ok(output) if output.status.success() => {
                let result = String::from_utf8_lossy(&output.stdout).trim().to_string();
                Some(result)
            }
            _ => None,
        }
    }
    
    #[cfg(target_os = "windows")]
    {
        // Windows would need win32 API or PowerShell
        eprintln!("[ColorDialog] Windows native color picker not yet implemented");
        None
    }
    
    #[cfg(not(any(target_os = "macos", target_os = "linux", target_os = "windows")))]
    {
        None
    }
}

/// Show native OS font picker and return selected font name and size
fn show_native_font_picker(default_font: &str, default_size: f64) -> Option<(String, f64)> {
    #[cfg(target_os = "macos")]
    {
        use std::process::Command;
        
        // Use AppleScript to show font selection from common fonts
        let fonts = vec![
            "Arial", "Helvetica", "Times New Roman", "Courier New", "Verdana",
            "Georgia", "Comic Sans MS", "Trebuchet MS", "Arial Black", "Impact",
            "Lucida Grande", "Monaco", "Menlo", "San Francisco", "System Font"
        ];
        
        let font_list = fonts.iter().map(|f| format!("\"{}\"", f)).collect::<Vec<_>>().join(", ");
        
        // Font name selection
        let script = format!(
            "choose from list {{{}}} with prompt \"Select Font:\" default items {{\"{}\"}}",
            font_list, default_font
        );
        
        let font_result = Command::new("osascript").arg("-e").arg(&script).output();
        
        if let Ok(output) = font_result {
            if output.status.success() {
                let selected_font = String::from_utf8_lossy(&output.stdout).trim().to_string();
                if selected_font != "false" {
                    // Now ask for size
                    let size_script = format!(
                        "display dialog \"Font Size:\" default answer \"{}\"",
                        default_size as i32
                    );
                    
                    if let Ok(size_output) = Command::new("osascript").arg("-e").arg(&size_script).output() {
                        if size_output.status.success() {
                            let result = String::from_utf8_lossy(&size_output.stdout);
                            if let Some(text_part) = result.split("text returned:").nth(1) {
                                if let Ok(size) = text_part.trim().parse::<f64>() {
                                    return Some((selected_font, size));
                                }
                            }
                        }
                    }
                }
            }
        }
        
        None
    }
    
    #[cfg(target_os = "linux")]
    {
        use std::process::Command;
        
        // Try yad or zenity for font selection
        match Command::new("yad")
            .arg("--font")
            .arg(format!("--fontname={} {}", default_font, default_size as i32))
            .output()
        {
            Ok(output) if output.status.success() => {
                let result = String::from_utf8_lossy(&output.stdout).trim().to_string();
                parse_font_string(&result)
            }
            _ => None,
        }
    }
    
    #[cfg(target_os = "windows")]
    {
        eprintln!("[FontDialog] Windows native font picker not yet implemented");
        None
    }
    
    #[cfg(not(any(target_os = "macos", target_os = "linux", target_os = "windows")))]
    {
        None
    }
}

/// Parse hex color to AppleScript RGB (0-65535)
fn parse_hex_to_applescript_rgb(hex: &str) -> (u32, u32, u32) {
    let hex = hex.trim_start_matches('#');
    if hex.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&hex[0..2], 16),
            u8::from_str_radix(&hex[2..4], 16),
            u8::from_str_radix(&hex[4..6], 16),
        ) {
            return (
                (r as u32) * 65535 / 255,
                (g as u32) * 65535 / 255,
                (b as u32) * 65535 / 255,
            );
        }
    }
    (0, 0, 0)
}

/// Parse AppleScript RGB output to hex color
fn applescript_rgb_to_hex(output: &str) -> Option<String> {
    let parts: Vec<&str> = output.trim().split(',').map(|s| s.trim()).collect();
    if parts.len() == 3 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            parts[0].parse::<u32>(),
            parts[1].parse::<u32>(),
            parts[2].parse::<u32>(),
        ) {
            let r8 = ((r * 255) / 65535) as u8;
            let g8 = ((g * 255) / 65535) as u8;
            let b8 = ((b * 255) / 65535) as u8;
            return Some(format!("#{:02X}{:02X}{:02X}", r8, g8, b8));
        }
    }
    None
}

/// Parse font string like "Arial 12" to (name, size)
#[allow(dead_code)]
fn parse_font_string(font_str: &str) -> Option<(String, f64)> {
    let parts: Vec<&str> = font_str.trim().rsplitn(2, ' ').collect();
    if parts.len() == 2 {
        if let Ok(size) = parts[0].parse::<f64>() {
            return Some((parts[1].to_string(), size));
        }
    }
    None
}
