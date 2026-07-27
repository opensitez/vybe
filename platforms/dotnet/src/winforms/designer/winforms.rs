use super::control::{Control, ControlType};
use super::errors::SaveResult;
use super::events::{EventBinding, EventType};
use super::form::Form;
use super::project::{FormFormat, FormModule};
use crate::winforms::designer::encoding::read_text_file;
use std::fs;
use std::path::Path;

struct CtrlInfo {
    name: String,
    control_type: ControlType,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    props: Vec<(String, String)>,
}

/// `"System.Windows.Forms.Button"` → `"Button"`
fn last_component(name: &str) -> &str {
    name.rsplit('.').next().unwrap_or(name)
}

pub fn load_form_vb(form_path: &Path) -> SaveResult<FormModule> {
    let user_code = read_text_file(form_path)?;

    let stem = form_path
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();
    let parent = form_path.parent().unwrap_or(Path::new("."));
    let designer_path = parent.join(format!("{}.Designer.vb", stem));

    let designer_code = if designer_path.exists() {
        read_text_file(&designer_path)?
    } else {
        String::new()
    };

    let form = parse_designer_fast(&designer_code, &user_code, &stem);
    Ok(FormModule::new_vbnet(form, designer_code, user_code))
}

// ── Fast line-based designer parser (no grammar / no AST) ────────────────

/// Parse designer + user code using pure string scanning.
/// Designer files follow a rigid pattern — no need for a full grammar parse.
fn parse_designer_fast(designer: &str, user_code: &str, fallback_name: &str) -> Form {
    let class_name = extract_class_name(designer)
        .or_else(|| extract_class_name(user_code))
        .unwrap_or_else(|| fallback_name.to_string());

    let mut form = Form::new(&class_name);

    // Collect control types from "Friend WithEvents X As Type" declarations
    // AND from "Me.X = New Type()" lines (the latter takes priority for type)
    let mut controls: Vec<CtrlInfo> = Vec::new();
    let mut in_init = false;

    for raw_line in designer.lines() {
        let line = raw_line.trim();

        // Skip comments and blank lines
        if line.is_empty() || line.starts_with('\'') {
            continue;
        }

        // Detect InitializeComponent boundaries
        if !in_init {
            let upper = line.to_uppercase();
            if upper.contains("SUB INITIALIZECOMPONENT") {
                in_init = true;
            }
            // Also pick up "Friend WithEvents X As Type" for type info
            if let Some(info) = parse_field_decl(line) {
                if let Some(ct) = vbnet_type_to_control_type(last_component(&info.1)) {
                    // Only add if not already known
                    if !controls
                        .iter()
                        .any(|c| c.name.eq_ignore_ascii_case(&info.0))
                    {
                        controls.push(CtrlInfo {
                            name: info.0,
                            control_type: ct,
                            x: 0,
                            y: 0,
                            width: 100,
                            height: 30,
                            props: Vec::new(),
                        });
                    }
                }
            }
            continue;
        }

        let upper = line.to_uppercase();
        if upper.starts_with("END SUB") {
            break;
        }

        // Me.X = New SomeType(...)
        if let Some((field, rhs)) = parse_me_assign(line) {
            if !field.contains('.') {
                // Simple field: Me.X = ...
                if let Some(type_name) = parse_new_expr(rhs) {
                    let short = last_component(&type_name);
                    if let Some(ct) = vbnet_type_to_control_type(short) {
                        // Update existing or insert new
                        if let Some(c) = controls
                            .iter_mut()
                            .find(|c| c.name.eq_ignore_ascii_case(field))
                        {
                            c.control_type = ct;
                        } else {
                            controls.push(CtrlInfo {
                                name: field.to_string(),
                                control_type: ct,
                                x: 0,
                                y: 0,
                                width: 100,
                                height: 30,
                                props: Vec::new(),
                            });
                        }
                        continue;
                    }
                }
                // Form-level property
                apply_form_property_str(&mut form, field, rhs);
            } else if let Some(dot) = field.find('.') {
                // Me.Ctrl.Prop = value
                let ctrl_name = &field[..dot];
                let prop_name = &field[dot + 1..];
                if let Some(c) = controls
                    .iter_mut()
                    .find(|c| c.name.eq_ignore_ascii_case(ctrl_name))
                {
                    apply_control_property(c, prop_name, rhs);
                }
            }
            continue;
        }

        // Me.Ctrl.DataBindings.Add("Prop", Me.Source, "Field")
        if let Some((ctrl, prop, source, field_col)) = parse_databindings_add(line) {
            if let Some(c) = controls
                .iter_mut()
                .find(|c| c.name.eq_ignore_ascii_case(ctrl))
            {
                c.props
                    .push(("DataBindings.Source".to_string(), source.to_string()));
                c.props
                    .push((format!("DataBindings.{}", prop), field_col.to_string()));
            }
            continue;
        }

        // Me.Ctrl.Items.AddRange(New String() { "a", "b" })
        if let Some((ctrl, items)) = parse_items_addrange(line) {
            if let Some(c) = controls
                .iter_mut()
                .find(|c| c.name.eq_ignore_ascii_case(ctrl))
            {
                c.props.push(("Items".to_string(), items.join("\n")));
            }
            continue;
        }

        // AddHandler Me.X.Event, AddressOf Me.Handler
        if let Some((ctrl, event, handler)) = parse_addhandler(line) {
            if let Some(et) = event_type_from_name(event) {
                form.event_bindings.push(EventBinding::with_handler(
                    ctrl.to_string(),
                    et,
                    handler.to_string(),
                ));
            }
        }
    }

    // Build Control objects
    for (tab_idx, ci) in controls.iter().enumerate() {
        let mut ctrl = Control::new(ci.control_type.clone(), ci.name.clone(), ci.x, ci.y);
        ctrl.bounds.width = ci.width;
        ctrl.bounds.height = ci.height;
        ctrl.tab_index = tab_idx as i32;
        for (prop, val) in &ci.props {
            ctrl.properties.set(prop.clone(), val.clone());
        }
        form.add_control(ctrl);
    }

    // Extract Handles clauses from user code
    collect_handles_from_source(user_code, &mut form);

    form
}

// ── Line parsers ─────────────────────────────────────────────────────────

/// Parse "Me.X = value" or "Me.X.Y = value", returns (field_path, rhs)
fn parse_me_assign<'a>(line: &'a str) -> Option<(&'a str, &'a str)> {
    let s = line.strip_prefix("Me.")?;
    let eq = s.find(" = ")?;
    let field = s[..eq].trim();
    let rhs = s[eq + 3..].trim();
    Some((field, rhs))
}

/// Parse "New System.Windows.Forms.TextBox(...)" → "System.Windows.Forms.TextBox"
fn parse_new_expr(rhs: &str) -> Option<&str> {
    let s = rhs.strip_prefix("New ")?;
    let end = s.find('(')?;
    Some(s[..end].trim())
}

/// Parse "Friend WithEvents X As System.Windows.Forms.TextBox" → ("X", "System.Windows.Forms.TextBox")
fn parse_field_decl(line: &str) -> Option<(String, String)> {
    let upper = line.to_uppercase();
    let idx = upper.find(" AS ")?;
    // Get the word before " As "
    let before = line[..idx].trim();
    let name = before.rsplit_once(' ')?.1.trim();
    let type_name = line[idx + 4..].trim();
    if name.is_empty() || type_name.is_empty() {
        return None;
    }
    Some((name.to_string(), type_name.to_string()))
}

/// Parse "AddHandler Me.Button1.Click, AddressOf Me.Button1_Click"
/// Returns (ctrl_name, event_name, handler_name)
fn parse_addhandler<'a>(line: &'a str) -> Option<(&'a str, &'a str, &'a str)> {
    let upper = line.to_uppercase();
    if !upper.starts_with("ADDHANDLER ") {
        return None;
    }
    let rest = line["AddHandler ".len()..].trim();
    let comma = rest.find(',')?;
    let event_part = rest[..comma].trim();
    let handler_part = rest[comma + 1..].trim();

    // event_part: "Me.Button1.Click" or "Me.Form1.Load"
    let event_path = event_part.strip_prefix("Me.").unwrap_or(event_part);
    let dot = event_path.rfind('.')?;
    let ctrl = &event_path[..dot];
    let event = &event_path[dot + 1..];

    // handler_part: "AddressOf Me.Button1_Click"
    let handler_upper = handler_part.to_uppercase();
    let handler = if let Some(pos) = handler_upper.find("ADDRESSOF ") {
        let after = handler_part[pos + "AddressOf ".len()..].trim();
        after.strip_prefix("Me.").unwrap_or(after)
    } else {
        handler_part
    };

    Some((ctrl, event, handler))
}

/// Parse `Me.Ctrl.DataBindings.Add("Prop", Me.Source, "Field")` →
/// `(ctrl, prop, source, field)`. Accepts both 3-arg and 4-arg forms
/// (the 4-arg takes a `True` for formatting_enabled which we discard).
fn parse_databindings_add<'a>(line: &'a str) -> Option<(&'a str, String, &'a str, String)> {
    let s = line.strip_prefix("Me.")?;
    let dot = s.find('.')?;
    let ctrl = &s[..dot];
    let after = &s[dot + 1..];
    // Must look like "DataBindings.Add(...)"
    let rest = after.strip_prefix("DataBindings.Add(")?;
    let close = rest.rfind(')')?;
    let args = &rest[..close];

    // Split by top-level commas (ignore commas inside quotes / parens).
    let mut parts: Vec<String> = Vec::new();
    let mut depth = 0i32;
    let mut in_str = false;
    let mut cur = String::new();
    for ch in args.chars() {
        match ch {
            '"' => {
                in_str = !in_str;
                cur.push(ch);
            }
            '(' if !in_str => {
                depth += 1;
                cur.push(ch);
            }
            ')' if !in_str => {
                depth -= 1;
                cur.push(ch);
            }
            ',' if !in_str && depth == 0 => {
                parts.push(std::mem::take(&mut cur).trim().to_string());
            }
            _ => cur.push(ch),
        }
    }
    if !cur.trim().is_empty() {
        parts.push(cur.trim().to_string());
    }
    if parts.len() < 3 {
        return None;
    }

    let prop = parse_string_value(&parts[0])?;
    // source is "Me.Name" or bare "Name"
    let src_raw = parts[1].trim();
    let source = src_raw.strip_prefix("Me.").unwrap_or(src_raw);
    let field_col = parse_string_value(&parts[2])?;

    // Static leak of the line's &str isn't possible — return owned
    // String for prop/field, but ctrl/source borrow from `line`.
    let source_static: &str = {
        // find source in original line for &str lifetime
        let pattern_me = format!("Me.{}", source);
        if let Some(p) = line.find(&pattern_me) {
            &line[p + 3..p + pattern_me.len()]
        } else if let Some(p) = line.find(source) {
            &line[p..p + source.len()]
        } else {
            return None;
        }
    };
    Some((ctrl, prop, source_static, field_col))
}

/// Parse `Me.Ctrl.Items.AddRange(New String() { "a", "b", "c" })` →
/// `(ctrl, [items])`. Also handles `Me.Ctrl.Items.Add("x")`.
fn parse_items_addrange<'a>(line: &'a str) -> Option<(&'a str, Vec<String>)> {
    let s = line.strip_prefix("Me.")?;
    let dot = s.find('.')?;
    let ctrl = &s[..dot];
    let after = &s[dot + 1..];
    // Match `Items.AddRange(…)` or `Items.Add(…)`
    let args = if let Some(r) = after.strip_prefix("Items.AddRange(") {
        let close = r.rfind(')')?;
        r[..close].to_string()
    } else if let Some(r) = after.strip_prefix("Items.Add(") {
        let close = r.rfind(')')?;
        let v = parse_string_value(&r[..close])?;
        return Some((ctrl, vec![v]));
    } else {
        return None;
    };
    // Drop the `New String() {` wrapper if present.
    let inner = args.trim();
    let inner = inner
        .trim_start_matches(|c: char| c != '{')
        .trim_start_matches('{')
        .trim_end_matches(|c: char| c != '}')
        .trim_end_matches('}')
        .trim();
    let mut items = Vec::new();
    let mut in_str = false;
    let mut cur = String::new();
    for ch in inner.chars() {
        match ch {
            '"' => {
                in_str = !in_str;
                cur.push(ch);
            }
            ',' if !in_str => {
                if let Some(v) = parse_string_value(cur.trim()) {
                    items.push(v);
                }
                cur.clear();
            }
            _ => cur.push(ch),
        }
    }
    if !cur.trim().is_empty() {
        if let Some(v) = parse_string_value(cur.trim()) {
            items.push(v);
        }
    }
    if items.is_empty() {
        return None;
    }
    Some((ctrl, items))
}

/// Extract `New Point(x, y)` or `New System.Drawing.Point(x, y)` → (x, y)
fn parse_point_value(rhs: &str) -> Option<(i32, i32)> {
    let type_name = parse_new_expr(rhs)?;
    let short = last_component(type_name);
    if !short.eq_ignore_ascii_case("Point") {
        return None;
    }
    parse_two_int_args(rhs)
}

/// Extract `New Size(w, h)` or `New System.Drawing.Size(w, h)` → (w, h)
fn parse_size_value(rhs: &str) -> Option<(i32, i32)> {
    let type_name = parse_new_expr(rhs)?;
    let short = last_component(type_name);
    if !short.eq_ignore_ascii_case("Size") && !short.eq_ignore_ascii_case("SizeF") {
        return None;
    }
    parse_two_int_args(rhs)
}

/// Extract two integer args from "New Type(a, b)"
fn parse_two_int_args(rhs: &str) -> Option<(i32, i32)> {
    let open = rhs.find('(')?;
    let close = rhs.rfind(')')?;
    let inner = rhs[open + 1..close].trim();
    let comma = inner.find(',')?;
    let a = parse_number(inner[..comma].trim())?;
    let b = parse_number(inner[comma + 1..].trim())?;
    Some((a, b))
}

/// Parse a number, handling "6.0!" VB single-precision suffix and float truncation
fn parse_number(s: &str) -> Option<i32> {
    let s = s
        .trim_end_matches('!')
        .trim_end_matches('F')
        .trim_end_matches('f');
    if let Ok(n) = s.parse::<i32>() {
        return Some(n);
    }
    if let Ok(f) = s.parse::<f64>() {
        return Some(f as i32);
    }
    None
}

/// Extract a VB string literal: "hello" → hello
fn parse_string_value(rhs: &str) -> Option<String> {
    let s = rhs.trim();
    if s.starts_with('"') && s.ends_with('"') && s.len() >= 2 {
        // Handle VB doubled quotes: "" → "
        let inner = &s[1..s.len() - 1];
        Some(inner.replace("\"\"", "\""))
    } else {
        None
    }
}

/// Apply a property to a control from string values
fn apply_control_property(c: &mut CtrlInfo, prop: &str, rhs: &str) {
    let prop_lower = prop.to_lowercase();
    match prop_lower.as_str() {
        "location" => {
            if let Some((x, y)) = parse_point_value(rhs) {
                c.x = x;
                c.y = y;
            }
        }
        "size" => {
            if let Some((w, h)) = parse_size_value(rhs) {
                c.width = w;
                c.height = h;
            }
        }
        "text" => {
            if let Some(s) = parse_string_value(rhs) {
                c.props.push(("Text".to_string(), s));
            }
        }
        "name" => { /* already have it */ }
        _ => {
            // Store the raw RHS value, cleaning up common patterns
            c.props.push((prop.to_string(), clean_value(rhs)));
        }
    }
}

/// Apply a form-level property from string values
fn apply_form_property_str(form: &mut Form, prop: &str, rhs: &str) {
    let prop_lower = prop.to_lowercase();
    match prop_lower.as_str() {
        "clientsize" | "size" => {
            if let Some((w, h)) = parse_size_value(rhs) {
                form.width = w;
                form.height = h;
            }
        }
        "text" => {
            if let Some(s) = parse_string_value(rhs) {
                form.text = s;
            }
        }
        "backcolor" => {
            form.back_color = Some(clean_value(rhs));
        }
        "forecolor" => {
            form.fore_color = Some(clean_value(rhs));
        }
        "name" => { /* already have it */ }
        _ => {
            form.properties.set(prop.to_string(), clean_value(rhs));
        }
    }
}

/// Clean up a raw RHS value for storage
fn clean_value(rhs: &str) -> String {
    let s = rhs.trim();
    // String literal
    if let Some(v) = parse_string_value(s) {
        return v;
    }
    // Boolean
    let upper = s.to_uppercase();
    if upper == "TRUE" {
        return "True".to_string();
    }
    if upper == "FALSE" {
        return "False".to_string();
    }
    // Everything else (enums, New expressions, etc.) — store as-is
    s.to_string()
}

/// Scan user code for "Sub X(...) Handles A.B, C.D" and wire event bindings
fn collect_handles_from_source(source: &str, form: &mut Form) {
    for line in source.lines() {
        let trimmed = line.trim();
        let upper = trimmed.to_uppercase();
        // Look for "Handles" keyword in Sub declarations
        if let Some(handles_pos) = upper.find(" HANDLES ") {
            // Extract handler name from "Sub X(" or "Sub X "
            let sub_name = extract_sub_name(trimmed);
            if sub_name.is_empty() {
                continue;
            }

            let handles_str = &trimmed[handles_pos + " Handles ".len()..];
            for target in handles_str.split(',') {
                let target = target.trim();
                if let Some(dot) = target.rfind('.') {
                    let ctrl = target[..dot].trim();
                    let event = target[dot + 1..].trim();
                    let ctrl = ctrl
                        .strip_prefix("Me.")
                        .or_else(|| ctrl.strip_prefix("MyBase."))
                        .unwrap_or(ctrl);
                    if let Some(et) = event_type_from_name(event) {
                        form.event_bindings.push(EventBinding::with_handler(
                            ctrl.to_string(),
                            et,
                            sub_name.to_string(),
                        ));
                    }
                }
            }
        }
    }
}

/// Extract the Sub/Function name from a declaration line
fn extract_sub_name(line: &str) -> &str {
    let upper = line.to_uppercase();
    let keyword = if let Some(pos) = upper.find(" SUB ") {
        pos + " SUB ".len()
    } else if let Some(pos) = upper.find(" FUNCTION ") {
        pos + " FUNCTION ".len()
    } else {
        return "";
    };
    let after = line[keyword..].trim();
    // Name ends at ( or space or end
    let end = after
        .find(|c: char| c == '(' || c.is_whitespace())
        .unwrap_or(after.len());
    &after[..end]
}

fn vbnet_type_to_control_type(name: &str) -> Option<ControlType> {
    match name.to_lowercase().as_str() {
        "button" => Some(ControlType::Button),
        "label" => Some(ControlType::Label),
        "textbox" => Some(ControlType::TextBox),
        "checkbox" => Some(ControlType::CheckBox),
        "radiobutton" => Some(ControlType::RadioButton),
        "combobox" => Some(ControlType::ComboBox),
        "listbox" => Some(ControlType::ListBox),
        "groupbox" => Some(ControlType::Frame),
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
        "timer" => Some(ControlType::Timer),
        "imagelist" => Some(ControlType::ImageList),
        "errorprovider" => Some(ControlType::ErrorProvider),
        "openfiledialog" => Some(ControlType::OpenFileDialog),
        "savefiledialog" => Some(ControlType::SaveFileDialog),
        "folderbrowserdialog" => Some(ControlType::FolderBrowserDialog),
        "fontdialog" => Some(ControlType::FontDialog),
        "colordialog" => Some(ControlType::ColorDialog),
        "printdialog" => Some(ControlType::PrintDialog),
        "notifyicon" => Some(ControlType::NotifyIcon),
        "checkedlistbox" => Some(ControlType::CheckedListBox),
        "domainupdown" => Some(ControlType::DomainUpDown),
        "propertygrid" => Some(ControlType::PropertyGrid),
        "splitter" => Some(ControlType::Splitter),
        "datagrid" => Some(ControlType::DataGrid),
        "usercontrol" => Some(ControlType::UserControl),
        // Data binding / ADO.NET components (WinForms classic names)
        "bindingsource" | "bindingsourcecomponent" => Some(ControlType::BindingSourceComponent),
        "dataset" | "datasetcomponent" => Some(ControlType::DataSetComponent),
        "datatable" | "datatablecomponent" => Some(ControlType::DataTableComponent),
        "dataadapter"
        | "sqldataadapter"
        | "oledbdataadapter"
        | "mysqldataadapter"
        | "odbcdataadapter"
        | "dataadaptercomponent" => Some(ControlType::DataAdapterComponent),
        "dataview" => Some(ControlType::DataView),
        "sqlconnection" => Some(ControlType::SqlConnection),
        "oledbconnection" => Some(ControlType::OleDbConnection),
        // Misc non-visual
        "helpprovider" => Some(ControlType::HelpProvider),
        "backgroundworker" => Some(ControlType::BackgroundWorker),
        // Print-related dialogs we missed
        "printdocument" => Some(ControlType::PrintDocument),
        "printpreviewdialog" => Some(ControlType::PrintPreviewDialog),
        "pagesetupdialog" => Some(ControlType::PageSetupDialog),
        _ => None,
    }
}

fn event_type_from_name(name: &str) -> Option<EventType> {
    match name.to_lowercase().as_str() {
        "click" => Some(EventType::Click),
        "dblclick" | "doubleclick" => Some(EventType::DoubleClick),
        "load" => Some(EventType::Load),
        "unload" => Some(EventType::Unload),
        "change" => Some(EventType::Change),
        "textchanged" => Some(EventType::TextChanged),
        "selectedindexchanged" => Some(EventType::SelectedIndexChanged),
        "checkedchanged" => Some(EventType::CheckedChanged),
        "valuechanged" => Some(EventType::ValueChanged),
        "keypress" => Some(EventType::KeyPress),
        "keydown" => Some(EventType::KeyDown),
        "keyup" => Some(EventType::KeyUp),
        "mousedown" => Some(EventType::MouseDown),
        "mouseup" => Some(EventType::MouseUp),
        "mousemove" => Some(EventType::MouseMove),
        "mouseenter" => Some(EventType::MouseEnter),
        "mouseleave" => Some(EventType::MouseLeave),
        "gotfocus" => Some(EventType::GotFocus),
        "lostfocus" => Some(EventType::LostFocus),
        "enter" => Some(EventType::Enter),
        "leave" => Some(EventType::Leave),
        "resize" => Some(EventType::Resize),
        "paint" => Some(EventType::Paint),
        "formclosing" => Some(EventType::FormClosing),
        "formclosed" => Some(EventType::FormClosed),
        "shown" => Some(EventType::Shown),
        "activated" => Some(EventType::Activated),
        "tick" => Some(EventType::Tick),
        "scroll" => Some(EventType::Scroll),
        "cellclick" => Some(EventType::CellClick),
        "linkclicked" => Some(EventType::LinkClicked),
        _ => None,
    }
}

/// Extract the class name from VB.NET source (looks for "Class <Name>")
fn extract_class_name(source: &str) -> Option<String> {
    for line in source.lines() {
        let upper = line.trim().to_uppercase();
        if upper.contains("CLASS ") {
            if let Some(pos) = upper.find("CLASS ") {
                let after = line.trim()[pos + 6..].trim();
                let name = after.split_whitespace().next().unwrap_or("");
                if !name.is_empty() {
                    return Some(name.to_string());
                }
            }
        }
    }
    None
}

pub fn save_form_vb(form_module: &FormModule, dir: &Path) -> SaveResult<()> {
    let name = &form_module.form.name;

    if let FormFormat::VbNet {
        designer_code,
        user_code,
    } = &form_module.format
    {
        let designer_path = dir.join(format!("{}.Designer.vb", name));
        fs::write(&designer_path, designer_code)?;

        let user_path = dir.join(format!("{}.vb", name));
        fs::write(&user_path, user_code)?;
    }

    Ok(())
}
