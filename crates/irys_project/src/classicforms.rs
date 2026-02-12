use crate::errors::SaveResult;
use crate::project::{FormFormat, FormModule};
use irys_forms::{Control, ControlType, Form};
use std::fs;
use std::path::Path;
use crate::encoding::read_text_file;

pub fn load_form_frm(path: &Path) -> SaveResult<FormModule> {
    let content = read_text_file(path)?;
    let lines: Vec<&str> = content.lines().collect();

    let mut form_name = String::new();
    let mut caption = String::new();
    let mut width = 0;
    let mut height = 0;
    let mut controls: Vec<Control> = Vec::new();

    let mut in_form = false;
    let mut in_control = false;
    let mut current_control: Option<Control> = None;
    let mut current_list_items: Vec<String> = Vec::new();
    let mut form_end_index: Option<usize> = None;

    let mut form_back_color: Option<String> = None;
    let mut form_fore_color: Option<String> = None;
    let mut form_font: Option<String> = None;

    for (idx, line) in lines.iter().enumerate() {
        let trimmed = line.trim();

        if !in_form {
            if trimmed.starts_with("Begin VB.Form") {
                form_name = trimmed.trim_start_matches("Begin VB.Form").trim().to_string();
                in_form = true;
            }
            continue;
        }

        if trimmed == "End" {
            if in_control {
                if let Some(mut control) = current_control.take() {
                    if !current_list_items.is_empty() {
                        control.set_list_items(current_list_items.clone());
                        current_list_items.clear();
                    }
                    // Post-parse fixup for legacy VB.Data controls:
                    // Infer the actual non-visual type from properties present.
                    if control.control_type == ControlType::DataAdapterComponent {
                        if control.properties.get_string("DataSource").is_some() {
                            control.control_type = ControlType::BindingSourceComponent;
                        } else if control.properties.get_string("DataSetName").is_some() {
                            control.control_type = ControlType::DataSetComponent;
                        } else if control.properties.get_string("TableName").is_some() {
                            control.control_type = ControlType::DataTableComponent;
                        }
                        // Otherwise stays DataAdapterComponent (has ConnectionString/SelectCommand or unknown)
                    }
                    controls.push(control);
                }
                in_control = false;
            } else {
                form_end_index = Some(idx);
                break;
            }
            continue;
        }

        if trimmed.starts_with("Begin ") && !trimmed.starts_with("Begin VB.Form") {
            let rest = trimmed.trim_start_matches("Begin ");
            let parts: Vec<&str> = rest.split_whitespace().collect();
            if parts.len() >= 2 {
                let vb_type = parts[0];
                let ctl_name = parts[1];
                let ctl_type = map_vb_type_to_control_type(vb_type);
                current_control = Some(Control::new(ctl_type, ctl_name.to_string(), 0, 0));
                current_list_items.clear();
                in_control = true;
            }
            continue;
        }

        if in_control {
            if let Some(control) = current_control.as_mut() {
                if trimmed.starts_with("Index") && !trimmed.starts_with("Index_") {
                    if let Some(val) = parse_prop_int(trimmed) { control.index = Some(val); }
                } else if trimmed.starts_with("Left") {
                    if let Some(val) = parse_prop_int(trimmed) { control.bounds.x = val; }
                } else if trimmed.starts_with("Top") {
                    if let Some(val) = parse_prop_int(trimmed) { control.bounds.y = val; }
                } else if trimmed.starts_with("Width") {
                    if let Some(val) = parse_prop_int(trimmed) { control.bounds.width = val; }
                } else if trimmed.starts_with("Height") {
                    if let Some(val) = parse_prop_int(trimmed) { control.bounds.height = val; }
                } else if trimmed.starts_with("Caption") {
                    if let Some(val) = parse_prop_string(trimmed) { control.set_caption(val); }
                } else if trimmed.starts_with("Text") {
                    if let Some(val) = parse_prop_string(trimmed) { control.set_text(val); }
                } else if trimmed.starts_with("BackColor") {
                    if let Some(val) = parse_prop_color(trimmed) { control.set_back_color(val); }
                } else if trimmed.starts_with("ForeColor") {
                    if let Some(val) = parse_prop_color(trimmed) { control.set_fore_color(val); }
                } else if trimmed.starts_with("Font") {
                    if let Some(val) = parse_prop_string(trimmed) { control.set_font(val); }
                } else if trimmed.starts_with("TabIndex") {
                    if let Some(val) = parse_prop_int(trimmed) { control.tab_index = val; }
                } else if trimmed.starts_with("Value") {
                    if let Some(val) = parse_prop_int(trimmed) {
                        use irys_forms::properties::PropertyValue;
                        control.properties.set_raw("Value", PropertyValue::Integer(val));
                    }
                } else if trimmed.starts_with("Enabled") {
                    if let Some(val) = parse_prop_bool(trimmed) { control.set_enabled(val); }
                } else if trimmed.starts_with("Visible") {
                    if let Some(val) = parse_prop_bool(trimmed) { control.set_visible(val); }
                } else if trimmed.starts_with("List") {
                    if let Some(val) = parse_prop_string(trimmed) { current_list_items.push(val); }
                } else if trimmed.starts_with("URL") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("URL", val); }
                } else if trimmed.starts_with("HTML") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("HTML", val); }
                } else if trimmed.starts_with("ToolbarVisible") {
                    if let Some(val) = parse_prop_bool(trimmed) { control.properties.set("ToolbarVisible", val); }
                } else if trimmed.starts_with("PathSeparator") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("PathSeparator", val); }
                } else if trimmed.starts_with("ConnectionString") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("ConnectionString", val); }
                } else if trimmed.starts_with("SelectCommand") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("SelectCommand", val); }
                } else if trimmed.starts_with("DataSource") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("DataSource", val); }
                } else if trimmed.starts_with("DataMember") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("DataMember", val); }
                } else if trimmed.starts_with("DataSetName") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("DataSetName", val); }
                } else if trimmed.starts_with("TableName") {
                    if let Some(val) = parse_prop_string(trimmed) { control.properties.set("TableName", val); }
                }
            }
        } else {
            if trimmed.starts_with("Caption") {
                if let Some(val) = parse_prop_string(trimmed) { caption = val; }
            } else if trimmed.starts_with("ClientWidth") || trimmed.starts_with("ScaleWidth") {
                if let Some(val) = parse_prop_int(trimmed) { width = val; }
            } else if trimmed.starts_with("ClientHeight") || trimmed.starts_with("ScaleHeight") {
                if let Some(val) = parse_prop_int(trimmed) { height = val; }
            } else if trimmed.starts_with("BackColor") {
                form_back_color = parse_prop_color(trimmed);
            } else if trimmed.starts_with("ForeColor") {
                form_fore_color = parse_prop_color(trimmed);
            } else if trimmed.starts_with("Font") {
                form_font = parse_prop_string(trimmed);
            }
        }
    }

    let mut real_code = String::new();
    if let Some(end_idx) = form_end_index {
        for line in lines.iter().skip(end_idx + 1) {
            if line.trim().starts_with("Attribute") {
                continue;
            }
            real_code.push_str(line);
            real_code.push('\n');
        }
    }

    let mut form = Form::new(&form_name);
    form.caption = caption;
    form.width = if width > 0 { width } else { 4800 };
    form.height = if height > 0 { height } else { 3600 };
    form.controls = controls;
    form.back_color = form_back_color;
    form.fore_color = form_fore_color;
    form.font = form_font;

    let form_name = form.name.clone();
    Ok(FormModule { form, code: real_code, format: FormFormat::Classic, resources: crate::resources::ResourceManager::new_named(form_name) })
}

fn map_vb_type_to_control_type(vb_type: &str) -> ControlType {
    match vb_type {
        "VB.CommandButton" => ControlType::Button,
        "VB.Label" => ControlType::Label,
        "VB.TextBox" => ControlType::TextBox,
        "VB.CheckBox" => ControlType::CheckBox,
        "VB.OptionButton" => ControlType::RadioButton,
        "VB.ComboBox" => ControlType::ComboBox,
        "VB.ListBox" => ControlType::ListBox,
        "VB.Frame" => ControlType::Frame,
        "VB.PictureBox" => ControlType::PictureBox,
        "RichTextLib.RichTextBox" => ControlType::RichTextBox,
        "SHDocVw.WebBrowser" => ControlType::WebBrowser,
        "MSComctlLib.TreeView" => ControlType::TreeView,
        "MSDataGridLib.DataGrid" => ControlType::DataGridView,
        "MSComctlLib.ListView" => ControlType::ListView,
        "VB.BindingSource" => ControlType::BindingSourceComponent,
        "VB.DataAdapter" => ControlType::DataAdapterComponent,
        "VB.DataSet" => ControlType::DataSetComponent,
        "VB.DataTable" => ControlType::DataTableComponent,
        "VB.Data" => ControlType::DataAdapterComponent, // legacy fallback
        _ => ControlType::Button,
    }
}

fn parse_prop_int(line: &str) -> Option<i32> {
    let parts: Vec<&str> = line.split('=').collect();
    if parts.len() > 1 {
        return parts[1].trim().parse::<i32>().ok();
    }
    None
}

fn parse_prop_string(line: &str) -> Option<String> {
    // Split only at the first '=' to preserve '=' chars inside the value
    // e.g. ConnectionString = "Server=localhost;Database=mydb" -> "Server=localhost;Database=mydb"
    if let Some(idx) = line.find('=') {
        let val = line[idx + 1..].trim().trim_matches('"');
        return Some(val.to_string());
    }
    None
}

fn parse_prop_bool(line: &str) -> Option<bool> {
    let parts: Vec<&str> = line.split('=').collect();
    if parts.len() > 1 {
        let raw = parts[1].trim();
        if raw.eq_ignore_ascii_case("true") || raw == "-1" { return Some(true); }
        if raw.eq_ignore_ascii_case("false") || raw == "0" { return Some(false); }
    }
    None
}

fn parse_prop_color(line: &str) -> Option<String> {
    let parts: Vec<&str> = line.split('=').collect();
    if parts.len() <= 1 { return None; }
    let raw = parts[1].trim();
    if raw.starts_with("&H") {
        let trimmed = raw.trim_start_matches("&H").trim_end_matches('&');
        if let Ok(val) = u32::from_str_radix(trimmed, 16) {
            // VB6 format is &H00BBGGRR& where RR is in bits 0-7, GG in 8-15, BB in 16-23
            let r = (val & 0xFF) as u8;
            let g = ((val >> 8) & 0xFF) as u8;
            let b = ((val >> 16) & 0xFF) as u8;
            return Some(format!("#{:02X}{:02X}{:02X}", r, g, b));
        }
    } else if raw.starts_with('#') {
        return Some(raw.to_string());
    }
    None
}
