use dioxus::prelude::*;
use crate::app_state::AppState;
use vybe_parser::{parse_program, Declaration};

fn handles_match(handle: &str, form_name: &str, control_name: &str, event_name: &str) -> bool {
    let parts: Vec<&str> = handle.split('.').collect();
    if parts.len() < 2 {
        return false;
    }

    let event_part = parts.last().unwrap();
    let control_part = parts.get(parts.len() - 2).unwrap();

    let resolved_control = if control_part.eq_ignore_ascii_case("me")
        || control_part.eq_ignore_ascii_case("mybase")
        || control_part.eq_ignore_ascii_case(form_name)
    {
        form_name
    } else {
        control_part
    };

    resolved_control.eq_ignore_ascii_case(control_name)
        && event_part.eq_ignore_ascii_case(event_name)
}

fn find_vbnet_handler(code: &str, form_name: &str, control_name: &str, event_name: &str) -> Option<String> {
    let program = parse_program(code).ok()?;
    let expected_name = format!("{}_{}", control_name, event_name);

    for decl in program.declarations {
        if let Declaration::Class(cls) = decl {
            if !cls.name.as_str().eq_ignore_ascii_case(form_name) {
                continue;
            }

            for method in cls.methods {
                if let vybe_parser::ast::decl::MethodDecl::Sub(sub) = method {
                    if sub.name.as_str().eq_ignore_ascii_case(&expected_name) {
                        return Some(sub.name.as_str().to_string());
                    }
                    if let Some(handles) = sub.handles.as_ref() {
                        if handles.iter().any(|h| handles_match(h, form_name, control_name, event_name)) {
                            return Some(sub.name.as_str().to_string());
                        }
                    }
                }
            }
        }
    }
    None
}

fn insert_before_end_class(code: &str, snippet: &str) -> String {
    let lower = code.to_lowercase();
    if let Some(idx) = lower.rfind("end class") {
        let (head, tail) = code.split_at(idx);
        let mut new_code = String::new();
        new_code.push_str(head.trim_end());
        new_code.push_str("\n\n");
        new_code.push_str(snippet);
        new_code.push_str("\n");
        new_code.push_str(tail);
        new_code
    } else {
        format!("{}\n\n{}", code, snippet)
    }
}

#[component]
pub fn PropertiesPanel() -> Element {
    let state = use_context::<AppState>();
    let selected_control_id = state.selected_controls.read().first().copied();
    let form_opt = state.get_current_form();
    
    let mut show_events = use_signal(|| false);
    let is_events = *show_events.read();
    
    let props_bg = if !is_events { "#e3f2fd" } else { "transparent" };
    let events_bg = if is_events { "#e3f2fd" } else { "transparent" };
    
    rsx! {
        div {
            class: "properties-panel",
            style: "width: 250px; background: #fafafa; border-left: 1px solid #ccc; padding: 8px; display: flex; flex-direction: column;",
            
            h3 { style: "margin: 0 0 8px 0; font-size: 14px;", "Properties" }
            
            // Tab switcher
            div {
                style: "display: flex; gap: 4px; margin-bottom: 8px; border-bottom: 1px solid #ccc; padding-bottom: 4px;",
                
                div {
                    style: "padding: 4px 12px; cursor: pointer; border-radius: 3px 3px 0 0; background: {props_bg};",
                    onclick: move |_| show_events.set(false),
                    "Properties"
                }
                
                div {
                    style: "padding: 4px 12px; cursor: pointer; border-radius: 3px 3px 0 0; background: {events_bg};",
                    onclick: move |_| show_events.set(true),
                    "Events"
                }
            }
            
            div {
                style: "flex: 1; overflow-y: auto; padding: 8px;",
                
                if is_events {
                     if let Some(cid) = selected_control_id {
                        if let Some(form) = form_opt.as_ref() {
                            if let Some(control) = form.get_control(cid) {
                                {
                                    let name = control.name.clone();
                                    rsx! {
                                        div {
                                            style: "font-weight: bold; margin-bottom: 8px;",
                                            "{name}"
                                        }
                                        
                                        {
                                            let is_array = control.is_array_member();
                                            let form_name = form.name.clone();
                                            let is_vbnet = state.is_current_form_vbnet();
                                            let events = match control.control_type {
                                                vybe_forms::ControlType::Button => vec!["Click", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "GotFocus", "LostFocus", "KeyDown", "KeyUp", "KeyPress", "EnabledChanged", "VisibleChanged", "Paint"],
                                                vybe_forms::ControlType::TextBox => vec!["TextChanged", "Change", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus", "Click", "MouseClick", "Enter", "Leave", "Validating", "Validated"],
                                                vybe_forms::ControlType::Label => vec!["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
                                                vybe_forms::ControlType::CheckBox => vec!["CheckedChanged", "Click", "GotFocus", "LostFocus", "KeyPress", "EnabledChanged"],
                                                vybe_forms::ControlType::RadioButton => vec!["CheckedChanged", "Click", "GotFocus", "LostFocus", "KeyPress", "EnabledChanged"],
                                                vybe_forms::ControlType::ListBox => vec!["SelectedIndexChanged", "Click", "DoubleClick", "MouseClick", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
                                                vybe_forms::ControlType::ComboBox => vec!["SelectedIndexChanged", "SelectedValueChanged", "TextChanged", "DropDown", "DropDownClosed", "Click", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
                                                vybe_forms::ControlType::Frame => vec!["Click", "DoubleClick", "MouseEnter", "MouseLeave"],
                                                vybe_forms::ControlType::Panel => vec!["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "MouseEnter", "MouseLeave", "Paint", "Scroll", "Resize"],
                                                vybe_forms::ControlType::PictureBox => vec!["Click", "DoubleClick", "MouseDown", "MouseUp", "MouseMove", "Paint"],
                                                vybe_forms::ControlType::RichTextBox => vec!["TextChanged", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus", "LinkClicked", "SelectionChanged"],
                                                vybe_forms::ControlType::TreeView => vec!["AfterSelect", "BeforeSelect", "AfterExpand", "AfterCollapse", "BeforeExpand", "BeforeCollapse", "NodeMouseClick", "NodeMouseDoubleClick", "AfterCheck", "BeforeCheck", "AfterLabelEdit", "BeforeLabelEdit", "ItemDrag", "Click", "DoubleClick", "KeyDown", "KeyPress"],
                                                vybe_forms::ControlType::ListView => vec!["SelectedIndexChanged", "ItemSelectionChanged", "ItemActivate", "ColumnClick", "ColumnWidthChanged", "ItemCheck", "Click", "DoubleClick", "MouseClick", "KeyDown", "KeyPress"],
                                                vybe_forms::ControlType::DataGridView => vec!["CellClick", "CellDoubleClick", "CellValueChanged", "CellContentClick", "CellEndEdit", "CellBeginEdit", "CellValidating", "CellEnter", "CellLeave", "CellFormatting", "SelectionChanged", "RowEnter", "RowLeave", "RowValidating", "RowValidated", "ColumnHeaderMouseClick", "CurrentCellChanged", "DataBindingComplete", "DataError", "Scroll", "KeyDown"],
                                                vybe_forms::ControlType::TabControl => vec!["SelectedIndexChanged", "Selected", "Deselecting", "Selecting", "Click", "DoubleClick"],
                                                vybe_forms::ControlType::ProgressBar => vec!["Click", "ValueChanged"],
                                                vybe_forms::ControlType::NumericUpDown => vec!["ValueChanged", "KeyPress", "KeyDown", "GotFocus", "LostFocus", "Enter", "Leave", "Validating"],
                                                vybe_forms::ControlType::MenuStrip | vybe_forms::ControlType::ContextMenuStrip => vec!["ItemClicked", "Click"],
                                                vybe_forms::ControlType::StatusStrip => vec!["ItemClicked", "Click"],
                                                vybe_forms::ControlType::ToolStripMenuItem => vec!["Click", "DropDownOpening", "DropDownClosed", "CheckedChanged"],
                                                vybe_forms::ControlType::ToolStripStatusLabel => vec!["Click", "DoubleClick"],
                                                vybe_forms::ControlType::WebBrowser => vec!["DocumentCompleted", "Navigating", "Navigated", "ProgressChanged", "Click"],
                                                vybe_forms::ControlType::DateTimePicker => vec!["ValueChanged", "DateChanged", "DropDown", "DropDownClosed", "GotFocus", "LostFocus", "KeyPress", "KeyDown"],
                                                vybe_forms::ControlType::LinkLabel => vec!["LinkClicked", "Click", "DoubleClick", "MouseEnter", "MouseLeave"],
                                                vybe_forms::ControlType::ToolStrip => vec!["ItemClicked", "ButtonClick", "Click"],
                                                vybe_forms::ControlType::TrackBar => vec!["Scroll", "ValueChanged", "MouseDown", "MouseUp", "GotFocus", "LostFocus"],
                                                vybe_forms::ControlType::MaskedTextBox => vec!["TextChanged", "MaskInputRejected", "KeyPress", "KeyDown", "KeyUp", "GotFocus", "LostFocus", "Enter", "Leave", "Validating"],
                                                vybe_forms::ControlType::SplitContainer => vec!["SplitterMoved", "SplitterMoving", "Click", "DoubleClick"],
                                                vybe_forms::ControlType::FlowLayoutPanel => vec!["Click", "DoubleClick", "Paint", "Resize", "Scroll"],
                                                vybe_forms::ControlType::TableLayoutPanel => vec!["Click", "DoubleClick", "Paint", "Resize", "Scroll", "CellPainting"],
                                                vybe_forms::ControlType::MonthCalendar => vec!["DateChanged", "DateSelected", "Click", "DoubleClick", "GotFocus", "LostFocus"],
                                                vybe_forms::ControlType::HScrollBar => vec!["Scroll", "ValueChanged", "GotFocus", "LostFocus"],
                                                vybe_forms::ControlType::VScrollBar => vec!["Scroll", "ValueChanged", "GotFocus", "LostFocus"],
                                                vybe_forms::ControlType::ToolTip => vec![],
                                                vybe_forms::ControlType::BindingNavigator => vec!["Click"],
                                                vybe_forms::ControlType::BindingSourceComponent => vec!["CurrentChanged", "PositionChanged", "DataSourceChanged"],
                                                vybe_forms::ControlType::DataSetComponent => vec![],
                                                vybe_forms::ControlType::DataTableComponent => vec!["RowChanged", "ColumnChanged"],
                                                vybe_forms::ControlType::DataAdapterComponent => vec!["FillError"],
                                                _ => vec!["Click", "DoubleClick"],
                                            };
                                            
                                            rsx! {
                                                for event_name in events {
                                                    {
                                                        let evt = event_name.to_string();
                                                        let c_name = name.clone();
                                                        let form_name = form_name.clone();
                                                        let mut state = state.clone();
                                                        let is_arr = is_array;
                                                        let is_vbnet_form = is_vbnet;
                                                        rsx! {
                                                            div {
                                                                key: "{event_name}",
                                                                style: "padding: 4px; border-bottom: 1px solid #eee; cursor: pointer;",
                                                                onclick: move |_| {
                                                                    // .NET-compatible event parameter signatures
                                                                    let params = if is_arr {
                                                                        "sender As Object, e As EventArgs, Index As Integer".to_string()
                                                                    } else {
                                                                        // Use EventType::parameters() from events.rs for correct signatures
                                                                        if let Some(event_type) = vybe_forms::EventType::from_name(&evt) {
                                                                            event_type.parameters().to_string()
                                                                        } else {
                                                                            // Fallback: determine EventArgs type based on event name
                                                                            let e_type = match evt.to_lowercase().as_str() {
                                                                                "mouseclick" | "mousedoubleclick" | "mousedown" | "mouseup" | "mousemove" | "mousewheel"
                                                                                | "nodemouseclick" | "nodemousedoubleclick" | "columnheadermouseclick" | "rowheadermouseclick" =>
                                                                                    "MouseEventArgs",
                                                                                "keydown" | "keyup" =>
                                                                                    "KeyEventArgs",
                                                                                "keypress" =>
                                                                                    "KeyPressEventArgs",
                                                                                "formclosing" =>
                                                                                    "FormClosingEventArgs",
                                                                                "formclosed" =>
                                                                                    "FormClosedEventArgs",
                                                                                "paint" | "cellpainting" =>
                                                                                    "PaintEventArgs",
                                                                                "cellclick" | "celldoubleclick" | "cellcontentclick" | "cellvaluechanged"
                                                                                | "cellendedit" | "cellbeginedit" | "cellvalidating" | "cellenter" | "cellleave"
                                                                                | "cellformatting" =>
                                                                                    "DataGridViewCellEventArgs",
                                                                                "dataerror" =>
                                                                                    "DataGridViewDataErrorEventArgs",
                                                                                "rowenter" | "rowleave" | "rowvalidating" | "rowvalidated" =>
                                                                                    "DataGridViewCellEventArgs",
                                                                                "afterselect" | "beforeselect" =>
                                                                                    "TreeViewEventArgs",
                                                                                "aftercheck" | "beforecheck" =>
                                                                                    "TreeViewEventArgs",
                                                                                "afterexpand" | "aftercollapse" | "beforeexpand" | "beforecollapse" =>
                                                                                    "TreeViewEventArgs",
                                                                                "afterlabeledit" | "beforelabeledit" =>
                                                                                    "NodeLabelEditEventArgs",
                                                                                "linkclicked" =>
                                                                                    "LinkLabelLinkClickedEventArgs",
                                                                                "splittermoved" | "splittermoving" =>
                                                                                    "SplitterEventArgs",
                                                                                "itemdrag" =>
                                                                                    "ItemDragEventArgs",
                                                                                "dragdrop" =>
                                                                                    "DragEventArgs",
                                                                                "dragenter" =>
                                                                                    "DragEventArgs",
                                                                                "dragover" =>
                                                                                    "DragEventArgs",
                                                                                "navigating" =>
                                                                                    "WebBrowserNavigatingEventArgs",
                                                                                "navigated" =>
                                                                                    "WebBrowserNavigatedEventArgs",
                                                                                "scroll" =>
                                                                                    "ScrollEventArgs",
                                                                                "drawitem" =>
                                                                                    "DrawItemEventArgs",
                                                                                "measureitem" =>
                                                                                    "MeasureItemEventArgs",
                                                                                "columnclick" =>
                                                                                    "ColumnClickEventArgs",
                                                                                "itemselectionchanged" =>
                                                                                    "ListViewItemSelectionChangedEventArgs",
                                                                                _ => "EventArgs",
                                                                            };
                                                                            format!("sender As Object, e As {}", e_type)
                                                                        }
                                                                    };
                                                                    let current_code = state.get_current_code();

                                                                    if is_vbnet_form {
                                                                        if find_vbnet_handler(&current_code, &form_name, &c_name, &evt).is_some() {
                                                                            state.show_code_editor.set(true);
                                                                            return;
                                                                        }

                                                                        let handler_name = format!("{}_{}", c_name, evt);
                                                                        let sub_decl = format!(
                                                                            "Private Sub {}({}) Handles {}.{}",
                                                                            handler_name, params, c_name, evt
                                                                        );

                                                                        if !current_code.contains(&sub_decl) {
                                                                            let new_code = insert_before_end_class(
                                                                                &current_code,
                                                                                &format!("{}\n    ' TODO: Add your code here\nEnd Sub", sub_decl),
                                                                            );
                                                                            state.update_current_code(new_code.clone());
                                                                            let escaped = new_code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                                                            let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                                                        }

                                                                        state.show_code_editor.set(true);
                                                                        return;
                                                                    }

                                                                    let handler_name = format!("{}_{}", c_name, evt);
                                                                    let sub_decl = format!("Private Sub {}({})", handler_name, params);

                                                                    if !current_code.contains(&sub_decl) {
                                                                        let new_code = format!("{}\n\n{}\n    ' TODO: Add your code here\nEnd Sub", current_code, sub_decl);
                                                                        state.update_current_code(new_code.clone());
                                                                        let escaped = new_code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                                                        let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                                                    }

                                                                    state.show_code_editor.set(true);
                                                                },
                                                                "{event_name}"
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                "Control not found"
                            }
                        } else {
                            "No form loaded"
                        }
                    } else if let Some(form) = form_opt.as_ref() {
                        {
                            let form_name = form.name.clone();
                            let is_vbnet = state.is_current_form_vbnet();
                            let events = vec!["Load", "Shown", "Activated", "Deactivate", "FormClosing", "FormClosed", "Resize", "Paint", "Click", "DoubleClick", "KeyDown", "KeyUp", "KeyPress", "MouseClick", "MouseDown", "MouseUp", "MouseMove"];
                            rsx! {
                                div { style: "font-weight: bold; margin-bottom: 8px;", "Form: {form_name}" }
                                for event_name in events {
                                    {
                                        let evt = event_name.to_string();
                                        let fname = form_name.clone();
                                        let mut state = state.clone();
                                        rsx! {
                                            div {
                                                key: "{event_name}",
                                                style: "padding: 4px; border-bottom: 1px solid #eee; cursor: pointer;",
                                                onclick: move |_| {
                                                    let params = if is_vbnet {
                                                        if let Some(et) = vybe_forms::EventType::from_name(&evt) {
                                                            et.parameters().to_string()
                                                        } else {
                                                            "sender As Object, e As EventArgs".to_string()
                                                        }
                                                    } else {
                                                        String::new()
                                                    };
                                                    let current_code = state.get_current_code();

                                                    if is_vbnet {
                                                        // Form-level handler uses Handles Me.Event
                                                        let handler_name = format!("{}_{}", fname, evt);
                                                        let sub_decl = format!(
                                                            "Private Sub {}({}) Handles Me.{}",
                                                            handler_name, params, evt
                                                        );

                                                        if !current_code.contains(&sub_decl) {
                                                            let new_code = insert_before_end_class(
                                                                &current_code,
                                                                &format!("{}\n    ' TODO: Add your code here\nEnd Sub", sub_decl),
                                                            );
                                                            state.update_current_code(new_code.clone());
                                                            let escaped = new_code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                                            let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                                        }

                                                        state.show_code_editor.set(true);
                                                        return;
                                                    }

                                                    let handler_name = format!("{}_{}", fname, evt);
                                                    let sub_decl = format!("Private Sub {}({})", handler_name, params);

                                                    if !current_code.contains(&sub_decl) {
                                                        let new_code = format!("{}\n\n{}\n    ' TODO: Add your code here\nEnd Sub", current_code, sub_decl);
                                                        state.update_current_code(new_code.clone());
                                                        let escaped = new_code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                                        let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                                    }

                                                    state.show_code_editor.set(true);
                                                },
                                                "{event_name}"
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        "No control selected"
                    }
                } else {
                    if let Some(cid) = selected_control_id {
                        if let Some(form) = form_opt.as_ref() {
                            if let Some(control) = form.get_control(cid) {
                                {
                                    let name = control.name.clone();
                                    let ctype = format!("{:?}", control.control_type);
                                    let x = control.bounds.x;
                                    let y = control.bounds.y;
                                    let w = control.bounds.width;
                                    let h = control.bounds.height;
                                    
                                    let caption = control.get_text().map(|s| s.to_string()).unwrap_or_default();
                                    let text = control.get_text().map(|s| s.to_string()).unwrap_or_default();
                                    let back_color = control.get_back_color().map(|s| s.to_string()).unwrap_or_else(|| "#f8fafc".to_string());
                                    let fore_color = control.get_fore_color().map(|s| s.to_string()).unwrap_or_else(|| "#0f172a".to_string());
                                    let font = control.get_font().map(|s| s.to_string()).unwrap_or_else(|| "Segoe UI, 12px".to_string());

                                    // Parse font into family + size
                                    let mut font_parts = font.split(',').map(|s| s.trim());
                                    let font_family = font_parts.next().unwrap_or("Segoe UI").to_string();
                                    let font_size_part = font_parts.next().unwrap_or("12px");
                                    let font_size_num: String = font_size_part.trim_end_matches("px").trim_end_matches("pt").to_string();
                                    let _font_family_sel = font_family.clone();
                                    let font_family_sel2 = font_family.clone();
                                    let font_size_sel = font_size_num.clone();
                                    let _font_size_sel2 = font_size_num.clone();
                                    
                                    let index_str = control.index.map(|i| i.to_string()).unwrap_or_default();
                                    let has_caption = matches!(control.control_type,
                                        vybe_forms::ControlType::Button |
                                        vybe_forms::ControlType::Label |
                                        vybe_forms::ControlType::CheckBox |
                                        vybe_forms::ControlType::RadioButton |
                                        vybe_forms::ControlType::Frame);
                                    let has_text = matches!(control.control_type, 
                                        vybe_forms::ControlType::TextBox | 
                                        vybe_forms::ControlType::ComboBox |
                                        vybe_forms::ControlType::MaskedTextBox |
                                        vybe_forms::ControlType::DateTimePicker |
                                        vybe_forms::ControlType::LinkLabel |
                                        vybe_forms::ControlType::ToolStripMenuItem |
                                        vybe_forms::ControlType::ToolStripStatusLabel |
                                        vybe_forms::ControlType::TabPage);
                                    let is_non_visual = control.control_type.is_non_visual();

                                    rsx! {
                                        div {
                                            style: "display: grid; grid-template-columns: 80px 1fr; gap: 4px; align-items: center;",
                                            
                                            div { style: "font-weight: bold;", "Name" }
                                            input {
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{name}",
                                                oninput: move |evt| {
                                                    state.update_control_property(cid, "Name", evt.value());
                                                }
                                            }

                                            div { style: "font-weight: bold;", "Index" }
                                            input {
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{index_str}",
                                                placeholder: "(none)",
                                                title: "Set a numeric index to make this a control array member",
                                                oninput: move |evt| {
                                                    state.update_control_property(cid, "Index", evt.value());
                                                }
                                            }

                                            div { style: "font-weight: bold;", "Type" }
                                            div { style: "font-size: 12px; color: #666;", "{ctype}" }
                                            
                                            if !is_non_visual {
                                            div { style: "font-weight: bold;", "Left" }
                                            input { 
                                                r#type: "number",
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{x}",
                                                oninput: move |evt| {
                                                    if let Ok(val) = evt.value().parse::<i32>() {
                                                        state.update_control_geometry(cid, val, y, w, h);
                                                    }
                                                }
                                            }
                                            
                                            div { style: "font-weight: bold;", "Top" }
                                            input { 
                                                r#type: "number",
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{y}",
                                                oninput: move |evt| {
                                                    if let Ok(val) = evt.value().parse::<i32>() {
                                                        state.update_control_geometry(cid, x, val, w, h);
                                                    }
                                                }
                                            }
                                            
                                            div { style: "font-weight: bold;", "Width" }
                                            input { 
                                                r#type: "number",
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{w}",
                                                oninput: move |evt| {
                                                    if let Ok(val) = evt.value().parse::<i32>() {
                                                        state.update_control_geometry(cid, x, y, val, h);
                                                    }
                                                }
                                            }
                                            
                                            div { style: "font-weight: bold;", "Height" }
                                            input { 
                                                r#type: "number",
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{h}",
                                                oninput: move |evt| {
                                                    if let Ok(val) = evt.value().parse::<i32>() {
                                                        state.update_control_geometry(cid, x, y, w, val);
                                                    }
                                                }
                                            }
                                            
                                            if has_caption {
                                                div { style: "font-weight: bold;", "Text" }
                                                input { 
                                                    style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                    value: "{caption}",
                                                    oninput: move |evt| {
                                                        state.update_control_property(cid, "Text", evt.value());
                                                    }
                                                }
                                            }

                                            // Appearance - BackColor
                                            div { style: "font-weight: bold;", "BackColor" }
                                            div { style: "display: flex; align-items: center; gap: 8px;", 
                                                input {
                                                    r#type: "color",
                                                    value: if back_color.starts_with('#') && back_color.len() == 7 { back_color.clone() } else { "#f8fafc".to_string() },
                                                    style: "width: 46px; height: 28px; padding: 0; border: 1px solid #ccc; background: transparent;",
                                                    onchange: move |evt| {
                                                        state.update_control_property(cid, "BackColor", evt.value());
                                                    }
                                                }
                                                input {
                                                    style: "flex: 1; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                    value: "{back_color}",
                                                    placeholder: "#RRGGBB or css color",
                                                    oninput: move |evt| {
                                                        state.update_control_property(cid, "BackColor", evt.value());
                                                    }
                                                }
                                            }

                                            // Appearance - ForeColor
                                            div { style: "font-weight: bold; margin-top: 6px;", "ForeColor" }
                                            div { style: "display: flex; align-items: center; gap: 8px;", 
                                                input {
                                                    r#type: "color",
                                                    value: if fore_color.starts_with('#') && fore_color.len() == 7 { fore_color.clone() } else { "#0f172a".to_string() },
                                                    style: "width: 46px; height: 28px; padding: 0; border: 1px solid #ccc; background: transparent;",
                                                    onchange: move |evt| {
                                                        state.update_control_property(cid, "ForeColor", evt.value());
                                                    }
                                                }
                                                input {
                                                    style: "flex: 1; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                    value: "{fore_color}",
                                                    placeholder: "#RRGGBB or css color",
                                                    oninput: move |evt| {
                                                        state.update_control_property(cid, "ForeColor", evt.value());
                                                    }
                                                }
                                            }

                                            // Appearance - Font (dropdowns)
                                            div { style: "font-weight: bold; margin-top: 6px;", "Font" }
                                            div { style: "display: flex; gap: 6px; align-items: center;", 
                                                select {
                                                    style: "flex: 1; border: 1px solid #ccc; padding: 4px; font-size: 12px;",
                                                    value: "{font_family}",
                                                    onchange: move |evt| {
                                                        let fam = evt.value();
                                                        let size = font_size_sel.clone();
                                                        state.update_control_property(cid, "Font", format!("{}, {}px", fam, size));
                                                    },
                                                    option { value: "Segoe UI", "Segoe UI" }
                                                    option { value: "Arial", "Arial" }
                                                    option { value: "Helvetica", "Helvetica" }
                                                    option { value: "Times New Roman", "Times New Roman" }
                                                    option { value: "Courier New", "Courier New" }
                                                    option { value: "Consolas", "Consolas" }
                                                    option { value: "Menlo", "Menlo" }
                                                    option { value: "Monaco", "Monaco" }
                                                    option { value: "Inter", "Inter" }
                                                    option { value: "Roboto", "Roboto" }
                                                }
                                                select {
                                                    style: "width: 90px; border: 1px solid #ccc; padding: 4px; font-size: 12px;",
                                                    value: "{font_size_num}",
                                                    onchange: move |evt| {
                                                        let size = evt.value();
                                                        let fam = font_family_sel2.clone();
                                                        state.update_control_property(cid, "Font", format!("{}, {}px", fam, size));
                                                    },
                                                    option { value: "10", "10" }
                                                    option { value: "11", "11" }
                                                    option { value: "12", "12" }
                                                    option { value: "14", "14" }
                                                    option { value: "16", "16" }
                                                    option { value: "18", "18" }
                                                    option { value: "20", "20" }
                                                }
                                            }
                                            

                                            
                                            if has_text {
                                                div { style: "font-weight: bold;", "Text" }
                                                input { 
                                                    style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                    value: "{text}",
                                                    oninput: move |evt| {
                                                        state.update_control_property(cid, "Text", evt.value());
                                                    }
                                                }
                                            }
                                            
                                            // Common properties for all controls
                                            div { style: "font-weight: bold;", "TabIndex" }
                                            input { 
                                                r#type: "number",
                                                style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{control.tab_index}",
                                                oninput: move |evt| {
                                                    state.update_control_property(cid, "TabIndex", evt.value());
                                                }
                                            }
                                            
                                            div { style: "font-weight: bold;", "Enabled" }
                                            input { 
                                                r#type: "checkbox",
                                                checked: control.is_enabled(),
                                                onchange: move |evt| {
                                                    state.update_control_property(cid, "Enabled", evt.checked().to_string());
                                                }
                                            }
                                            
                                            div { style: "font-weight: bold;", "Visible" }
                                            input { 
                                                r#type: "checkbox",
                                                checked: control.is_visible(),
                                                onchange: move |evt| {
                                                    state.update_control_property(cid, "Visible", evt.checked().to_string());
                                                }
                                            }
                                            } // end if !is_non_visual
                                            
                                            // ---- Data Binding section ----
                                            // ALL controls get data binding options
                                            {
                                                {
                                                    // Collect available data components on the form (exclude self!)
                                                    let control_id = control.id;
                                                    let form_opt2 = state.get_current_form();
                                                    // For visual controls: they bind to a BindingSource
                                                    let binding_sources: Vec<String> = form_opt2.as_ref()
                                                        .map(|f| f.controls.iter()
                                                            .filter(|c| c.id != control_id && matches!(c.control_type,
                                                                vybe_forms::ControlType::BindingSourceComponent))
                                                            .map(|c| c.name.clone())
                                                            .collect())
                                                        .unwrap_or_default();
                                                    // For BindingSource: its DataSource can be DataAdapter, DataSet, or DataTable
                                                    let bs_data_sources: Vec<String> = form_opt2.as_ref()
                                                        .map(|f| f.controls.iter()
                                                            .filter(|c| c.id != control_id && matches!(c.control_type,
                                                                vybe_forms::ControlType::DataAdapterComponent |
                                                                vybe_forms::ControlType::DataSetComponent |
                                                                vybe_forms::ControlType::DataTableComponent))
                                                            .map(|c| c.name.clone())
                                                            .collect())
                                                        .unwrap_or_default();
                                                    let has_complex_binding = control.control_type.supports_complex_binding();

                                                    // Helper: resolve a connection string's relative Data Source path
                                                    // against the project directory so the editor can find the DB file.
                                                    let project_dir: Option<std::path::PathBuf> = state.current_project_path.read()
                                                        .as_ref()
                                                        .and_then(|p| p.parent().map(|d| d.to_path_buf()));
                                                    let resolve_conn_str = move |conn_str: &str| -> String {
                                                        if let Some(ref dir) = project_dir {
                                                            // Parse "Data Source=relative.db" and make absolute
                                                            let lower = conn_str.to_lowercase();
                                                            if let Some(pos) = lower.find("data source=") {
                                                                let start = pos + 12;
                                                                let rest = &conn_str[start..];
                                                                let end = rest.find(';').unwrap_or(rest.len());
                                                                let db_path = rest[..end].trim();
                                                                if !db_path.is_empty() && db_path != ":memory:" && !std::path::Path::new(db_path).is_absolute() {
                                                                    let abs = dir.join(db_path);
                                                                    let resolved = format!("{}Data Source={}{}",
                                                                        &conn_str[..pos],
                                                                        abs.display(),
                                                                        &rest[end..]);
                                                                    return resolved;
                                                                }
                                                            }
                                                        }
                                                        conn_str.to_string()
                                                    };

                                                    // Helper: resolve available column names by walking the binding chain
                                                    // control -> BindingSource (bs_name) -> DataAdapter -> ConnectionString + SelectCommand -> columns
                                                    // Also considers BindingSource.DataMember (table name) as fallback query
                                                    let resolve_columns_for_bs = |bs_name: &str| -> Vec<String> {
                                                        if bs_name.is_empty() {
                                                            eprintln!("[resolve_columns] bs_name is empty");
                                                            return Vec::new();
                                                        }
                                                        let form_ref = state.get_current_form();
                                                        let form = match form_ref.as_ref() {
                                                            Some(f) => f,
                                                            None => { eprintln!("[resolve_columns] no current form"); return Vec::new(); }
                                                        };
                                                        // Find the BindingSource control
                                                        let bs_ctrl = form.controls.iter()
                                                            .find(|c| c.name.eq_ignore_ascii_case(bs_name)
                                                                && matches!(c.control_type, vybe_forms::ControlType::BindingSourceComponent));
                                                        let bs_ctrl = match bs_ctrl {
                                                            Some(c) => c,
                                                            None => {
                                                                eprintln!("[resolve_columns] BindingSource '{}' not found among {} controls: {:?}",
                                                                    bs_name, form.controls.len(),
                                                                    form.controls.iter().map(|c| format!("{}({:?})", c.name, c.control_type)).collect::<Vec<_>>());
                                                                return Vec::new();
                                                            }
                                                        };
                                                        // Get the DataAdapter name from the BindingSource
                                                        let da_name = match bs_ctrl.properties.get_string("DataSource") {
                                                            Some(s) if !s.is_empty() => s.to_string(),
                                                            _ => {
                                                                eprintln!("[resolve_columns] BindingSource '{}' has no DataSource property. Props: {:?}",
                                                                    bs_name, bs_ctrl.properties.iter().map(|(k,v)| format!("{}={:?}", k, v)).collect::<Vec<_>>());
                                                                return Vec::new();
                                                            }
                                                        };
                                                        // Get the DataMember (table name) from the BindingSource
                                                        let data_member = bs_ctrl.properties.get_string("DataMember")
                                                            .map(|s| s.to_string()).unwrap_or_default();
                                                        eprintln!("[resolve_columns] BS '{}' -> DataAdapter '{}', DataMember '{}'", bs_name, da_name, data_member);
                                                        // Find the DataAdapter control
                                                        let da_ctrl = form.controls.iter()
                                                            .find(|c| c.name.eq_ignore_ascii_case(&da_name)
                                                                && matches!(c.control_type, vybe_forms::ControlType::DataAdapterComponent));
                                                        let da_ctrl = match da_ctrl {
                                                            Some(c) => c,
                                                            None => {
                                                                eprintln!("[resolve_columns] DataAdapter '{}' not found", da_name);
                                                                return Vec::new();
                                                            }
                                                        };
                                                        let conn_str = da_ctrl.properties.get_string("ConnectionString").unwrap_or("");
                                                        if conn_str.is_empty() {
                                                            eprintln!("[resolve_columns] DataAdapter '{}' has empty ConnectionString", da_name);
                                                            return Vec::new();
                                                        }
                                                        let conn_str = resolve_conn_str(conn_str);
                                                        // Use the DataAdapter's SelectCommand if available,
                                                        // otherwise fall back to "SELECT * FROM <DataMember>" if a table is selected
                                                        let da_select = da_ctrl.properties.get_string("SelectCommand").unwrap_or("").to_string();
                                                        let query = if !da_select.is_empty() {
                                                            da_select
                                                        } else if !data_member.is_empty() {
                                                            format!("SELECT * FROM {}", data_member)
                                                        } else {
                                                            eprintln!("[resolve_columns] No SelectCommand and no DataMember for DA '{}'", da_name);
                                                            return Vec::new();
                                                        };
                                                        eprintln!("[resolve_columns] conn='{}', query='{}'", conn_str, query);
                                                        match vybe_runtime::data_access::fetch_columns_for_query(&conn_str, &query) {
                                                            Ok(cols) => {
                                                                eprintln!("[resolve_columns] SUCCESS: {:?}", cols);
                                                                cols
                                                            }
                                                            Err(e) => {
                                                                eprintln!("[resolve_columns] ERROR: {}", e);
                                                                Vec::new()
                                                            }
                                                        }
                                                    };

                                                    eprintln!("[DATA_SECTION] Rendering for '{}' ({:?}), is_non_visual={}, has_complex={}, binding_sources={:?}",
                                                        control.name, control.control_type, is_non_visual, has_complex_binding, binding_sources);

                                                    rsx! {
                                                        div { style: "grid-column: 1 / -1; margin-top: 8px; padding-top: 6px; border-top: 1px solid #ddd; font-weight: bold; font-size: 11px; color: #0078d4; text-transform: uppercase;",
                                                            "Data"
                                                        }

                                                        // === Simple Data Bindings for visual controls WITHOUT complex binding ===
                                                        // (TextBox→Text, Label→Text, CheckBox→Checked, etc.)
                                                        // Controls with complex binding (DataGridView, ListBox, ComboBox) use DataSource instead
                                                        if !is_non_visual && !has_complex_binding {
                                                            {
                                                                // Determine which control property to bind (the "bindable property")
                                                                let bindable_prop = match control.control_type {
                                                                    vybe_forms::ControlType::TextBox | vybe_forms::ControlType::RichTextBox => "Text",
                                                                    vybe_forms::ControlType::Label => "Text",
                                                                    vybe_forms::ControlType::CheckBox | vybe_forms::ControlType::RadioButton => "Checked",
                                                                    vybe_forms::ControlType::Button => "Text",
                                                                    vybe_forms::ControlType::PictureBox => "ImageLocation",
                                                                    _ => "Text",
                                                                };
                                                                let binding_key = format!("DataBindings.{}", bindable_prop);
                                                                let current_binding_bs = control.properties.get_string("DataBindings.Source")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let current_binding_col = control.properties.get_string(&binding_key)
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let columns = resolve_columns_for_bs(&current_binding_bs);

                                                                rsx! {
                                                                    div { style: "font-weight: bold; font-size: 11px;", "DataSource" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DataBindings.Source", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_binding_bs.is_empty(), "(none)" }
                                                                        for bs_name in &binding_sources {
                                                                            option {
                                                                                value: "{bs_name}",
                                                                                selected: current_binding_bs == *bs_name,
                                                                                "{bs_name}"
                                                                            }
                                                                        }
                                                                    }

                                                                    div { style: "font-weight: bold; font-size: 11px;", "{bindable_prop}" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        title: "DataBindings.Add(\"{bindable_prop}\", bindingSource, \"ColumnName\")",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, &binding_key, evt.value());
                                                                        },
                                                                        option { value: "", selected: current_binding_col.is_empty(), "(none)" }
                                                                        for col in &columns {
                                                                            option {
                                                                                value: "{col}",
                                                                                selected: current_binding_col == *col,
                                                                                "{col}"
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // === Complex binding (DataSource property) for list/grid controls ===
                                                        // DataGridView, ListBox, ComboBox bind to a BindingSource via DataSource property
                                                        if has_complex_binding && !matches!(control.control_type, vybe_forms::ControlType::BindingNavigator) {
                                                            {
                                                                let current_ds = control.properties.get_string("DataSource")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let current_dm = control.properties.get_string("DataMember")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let columns = resolve_columns_for_bs(&current_ds);
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "DataSource" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DataSource", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_ds.is_empty(), "(none)" }
                                                                        for bs_name in &binding_sources {
                                                                            option {
                                                                                value: "{bs_name}",
                                                                                selected: current_ds == *bs_name,
                                                                                "{bs_name}"
                                                                            }
                                                                        }
                                                                    }

                                                                    div { style: "font-weight: bold;", "DataMember" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DataMember", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_dm.is_empty(), "(none)" }
                                                                        for col in &columns {
                                                                            option {
                                                                                value: "{col}",
                                                                                selected: current_dm == *col,
                                                                                "{col}"
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // === BindingSource DataSource: binds to DataAdapter/DataSet/DataTable ===
                                                        if matches!(control.control_type, vybe_forms::ControlType::BindingSourceComponent) {
                                                            {
                                                                let current_ds = control.properties.get_string("DataSource")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let current_dm = control.properties.get_string("DataMember")
                                                                    .map(|s| s.to_string()).unwrap_or_default();

                                                                // Resolve tables for DataMember dropdown:
                                                                // If the DataSource is a DataAdapter, we can list tables from its connection
                                                                let da_tables: Vec<String> = if !current_ds.is_empty() {
                                                                    let form_ref = state.get_current_form();
                                                                    form_ref.as_ref().and_then(|f| {
                                                                        let da = f.controls.iter()
                                                                            .find(|c| c.name.eq_ignore_ascii_case(&current_ds)
                                                                                && matches!(c.control_type, vybe_forms::ControlType::DataAdapterComponent))?;
                                                                        let cs = da.properties.get_string("ConnectionString")?;
                                                                        if cs.is_empty() { return None; }
                                                                        let cs = resolve_conn_str(cs);
                                                                        vybe_runtime::data_access::test_connection_and_list_tables(&cs).ok()
                                                                    }).unwrap_or_default()
                                                                } else {
                                                                    Vec::new()
                                                                };

                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "DataSource" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DataSource", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_ds.is_empty(), "(none)" }
                                                                        for ds_name in &bs_data_sources {
                                                                            option {
                                                                                value: "{ds_name}",
                                                                                selected: current_ds == *ds_name,
                                                                                "{ds_name}"
                                                                            }
                                                                        }
                                                                    }

                                                                    div { style: "font-weight: bold;", "DataMember" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DataMember", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_dm.is_empty(), "(none)" }
                                                                        for tbl in &da_tables {
                                                                            option {
                                                                                value: "{tbl}",
                                                                                selected: current_dm == *tbl,
                                                                                "{tbl}"
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // DisplayMember/ValueMember for ListBox/ComboBox
                                                        if matches!(control.control_type, vybe_forms::ControlType::ListBox | vybe_forms::ControlType::ComboBox) {
                                                            {
                                                                let current_ds = control.properties.get_string("DataSource")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let columns = resolve_columns_for_bs(&current_ds);
                                                                let display_member = control.properties.get_string("DisplayMember")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let value_member = control.properties.get_string("ValueMember")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "DisplayMember" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "DisplayMember", evt.value());
                                                                        },
                                                                        option { value: "", selected: display_member.is_empty(), "(none)" }
                                                                        for col in &columns {
                                                                            option {
                                                                                value: "{col}",
                                                                                selected: display_member == *col,
                                                                                "{col}"
                                                                            }
                                                                        }
                                                                    }
                                                                    div { style: "font-weight: bold;", "ValueMember" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "ValueMember", evt.value());
                                                                        },
                                                                        option { value: "", selected: value_member.is_empty(), "(none)" }
                                                                        for col in &columns {
                                                                            option {
                                                                                value: "{col}",
                                                                                selected: value_member == *col,
                                                                                "{col}"
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // BindingNavigator: link to BindingSource
                                                        if matches!(control.control_type, vybe_forms::ControlType::BindingNavigator) {
                                                            {
                                                                let current_bs = control.properties.get_string("BindingSource")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "BindingSource" }
                                                                    select {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        onchange: move |evt| {
                                                                            state.update_control_property(cid, "BindingSource", evt.value());
                                                                        },
                                                                        option { value: "", selected: current_bs.is_empty(), "(none)" }
                                                                        for bs_name in &binding_sources {
                                                                            option {
                                                                                value: "{bs_name}",
                                                                                selected: current_bs == *bs_name,
                                                                                "{bs_name}"
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // BindingSource-specific properties
                                                        if matches!(control.control_type, vybe_forms::ControlType::BindingSourceComponent) {
                                                            {
                                                                let filter = control.properties.get_string("Filter")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let sort = control.properties.get_string("Sort")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "Filter" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{filter}",
                                                                        placeholder: "e.g. Name = 'Test'",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "Filter", evt.value());
                                                                        }
                                                                    }
                                                                    div { style: "font-weight: bold;", "Sort" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{sort}",
                                                                        placeholder: "e.g. Name ASC",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "Sort", evt.value());
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // DataSet-specific properties
                                                        if matches!(control.control_type, vybe_forms::ControlType::DataSetComponent) {
                                                            {
                                                                let dsn = control.properties.get_string("DataSetName")
                                                                    .map(|s| s.to_string()).unwrap_or_else(|| "NewDataSet".to_string());
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "DataSetName" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{dsn}",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "DataSetName", evt.value());
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // DataTable-specific properties
                                                        if matches!(control.control_type, vybe_forms::ControlType::DataTableComponent) {
                                                            {
                                                                let tn = control.properties.get_string("TableName")
                                                                    .map(|s| s.to_string()).unwrap_or_else(|| "Table1".to_string());
                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "TableName" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{tn}",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "TableName", evt.value());
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // DataAdapter-specific properties (no DataSource — it IS the source)
                                                        if matches!(control.control_type, vybe_forms::ControlType::DataAdapterComponent) {
                                                            {
                                                                let sc = control.properties.get_string("SelectCommand")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let cs = control.properties.get_string("ConnectionString")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let db_type = control.properties.get_string("DbType")
                                                                    .map(|s| s.to_string()).unwrap_or_else(|| "SQLite".to_string());
                                                                let db_path = control.properties.get_string("DbPath")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let db_host = control.properties.get_string("DbHost")
                                                                    .map(|s| s.to_string()).unwrap_or_else(|| "localhost".to_string());
                                                                let db_port = control.properties.get_string("DbPort")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let db_name = control.properties.get_string("DbName")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let db_user = control.properties.get_string("DbUser")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let db_pass = control.properties.get_string("DbPassword")
                                                                    .map(|s| s.to_string()).unwrap_or_default();
                                                                let mut show_conn_builder = use_signal(|| false);
                                                                let is_builder_open = *show_conn_builder.read();
                                                                let mut table_list: Signal<Vec<String>> = use_signal(|| Vec::new());
                                                                let mut conn_status = use_signal(|| String::new());
                                                                let tables = table_list.read().clone();

                                                                rsx! {
                                                                    div { style: "font-weight: bold;", "SelectCmd" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{sc}",
                                                                        placeholder: "SELECT * FROM ...",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "SelectCommand", evt.value());
                                                                        }
                                                                    }

                                                                    div { style: "font-weight: bold;", "ConnStr" }
                                                                    input {
                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                        value: "{cs}",
                                                                        placeholder: "Data Source=...",
                                                                        oninput: move |evt| {
                                                                            state.update_control_property(cid, "ConnectionString", evt.value());
                                                                        }
                                                                    }

                                                                    button {
                                                                        style: "grid-column: 1 / -1; margin-top: 2px; width: 100%; padding: 4px 8px; border: 1px solid #0078d4; background: #0078d4; color: white; cursor: pointer; border-radius: 3px; font-size: 11px;",
                                                                        onclick: move |_| {
                                                                            show_conn_builder.set(!is_builder_open);
                                                                        },
                                                                        if is_builder_open { "▲ Hide Connection Builder" } else { "🔧 Connection String Builder..." }
                                                                    }

                                                                    // Connection String Builder panel
                                                                    if is_builder_open {
                                                                        div { style: "grid-column: 1 / -1; background: #f5f9ff; border: 1px solid #a0c4e8; border-radius: 4px; padding: 8px; margin-top: 4px;",
                                                                            div { style: "font-weight: bold; font-size: 11px; color: #0078d4; margin-bottom: 6px;", "🔧 Connection Builder" }

                                                                            div { style: "display: grid; grid-template-columns: 70px 1fr; gap: 4px; align-items: center;",
                                                                                div { style: "font-size: 11px; font-weight: bold;", "Server" }
                                                                                select {
                                                                                    style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                    value: "{db_type}",
                                                                                    onchange: move |evt| {
                                                                                        let val = evt.value();
                                                                                        state.update_control_property(cid, "DbType", val.clone());
                                                                                        // Set default port
                                                                                        let port = match val.as_str() {
                                                                                            "PostgreSQL" => "5432",
                                                                                            "MySQL" => "3306",
                                                                                            _ => "",
                                                                                        };
                                                                                        state.update_control_property(cid, "DbPort", port.to_string());
                                                                                    },
                                                                                    option { value: "SQLite", "SQLite" }
                                                                                    option { value: "PostgreSQL", "PostgreSQL" }
                                                                                    option { value: "MySQL", "MySQL" }
                                                                                }

                                                                                if db_type == "SQLite" {
                                                                                    div { style: "font-size: 11px; font-weight: bold;", "File" }
                                                                                    input {
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_path}",
                                                                                        placeholder: "database.db  or  :memory:",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbPath", evt.value());
                                                                                        }
                                                                                    }
                                                                                } else {
                                                                                    div { style: "font-size: 11px; font-weight: bold;", "Host" }
                                                                                    input {
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_host}",
                                                                                        placeholder: "localhost",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbHost", evt.value());
                                                                                        }
                                                                                    }

                                                                                    div { style: "font-size: 11px; font-weight: bold;", "Port" }
                                                                                    input {
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_port}",
                                                                                        placeholder: "5432",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbPort", evt.value());
                                                                                        }
                                                                                    }

                                                                                    div { style: "font-size: 11px; font-weight: bold;", "Database" }
                                                                                    input {
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_name}",
                                                                                        placeholder: "mydb",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbName", evt.value());
                                                                                        }
                                                                                    }

                                                                                    div { style: "font-size: 11px; font-weight: bold;", "User" }
                                                                                    input {
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_user}",
                                                                                        placeholder: "username",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbUser", evt.value());
                                                                                        }
                                                                                    }

                                                                                    div { style: "font-size: 11px; font-weight: bold;", "Password" }
                                                                                    input {
                                                                                        r#type: "password",
                                                                                        style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                                        value: "{db_pass}",
                                                                                        placeholder: "••••••",
                                                                                        oninput: move |evt| {
                                                                                            state.update_control_property(cid, "DbPassword", evt.value());
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }

                                                                            button {
                                                                                style: "margin-top: 8px; width: 100%; padding: 4px 8px; border: 1px solid #0078d4; background: #0078d4; color: white; cursor: pointer; border-radius: 3px; font-size: 11px;",
                                                                                onclick: move |_| {
                                                                                    // Build connection string from fields
                                                                                    let form_opt3 = state.get_current_form();
                                                                                    if let Some(form) = form_opt3.as_ref() {
                                                                                        if let Some(ctrl) = form.get_control(cid) {
                                                                                            let dtype = ctrl.properties.get_string("DbType").unwrap_or("SQLite");
                                                                                            let conn = match dtype {
                                                                                                "SQLite" => {
                                                                                                    let path = ctrl.properties.get_string("DbPath").unwrap_or("database.db");
                                                                                                    format!("Data Source={}", path)
                                                                                                }
                                                                                                "PostgreSQL" => {
                                                                                                    let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                                                                                    let port = ctrl.properties.get_string("DbPort").unwrap_or("5432");
                                                                                                    let db = ctrl.properties.get_string("DbName").unwrap_or("mydb");
                                                                                                    let user = ctrl.properties.get_string("DbUser").unwrap_or("postgres");
                                                                                                    let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                                                                                    format!("Host={};Port={};Database={};Username={};Password={}", host, port, db, user, pass)
                                                                                                }
                                                                                                "MySQL" => {
                                                                                                    let host = ctrl.properties.get_string("DbHost").unwrap_or("localhost");
                                                                                                    let port = ctrl.properties.get_string("DbPort").unwrap_or("3306");
                                                                                                    let db = ctrl.properties.get_string("DbName").unwrap_or("mydb");
                                                                                                    let user = ctrl.properties.get_string("DbUser").unwrap_or("root");
                                                                                                    let pass = ctrl.properties.get_string("DbPassword").unwrap_or("");
                                                                                                    format!("Server={};Port={};Database={};Uid={};Pwd={}", host, port, db, user, pass)
                                                                                                }
                                                                                                _ => String::new(),
                                                                                            };
                                                                                            state.update_control_property(cid, "ConnectionString", conn);
                                                                                        }
                                                                                    }
                                                                                    show_conn_builder.set(false);
                                                                                },
                                                                                "Build Connection String"
                                                                            }
                                                                        }
                                                                    }

                                                                    // Test Connection & Table picker
                                                                    {
                                                                        let status_text = conn_status.read().clone();
                                                                        let status_bg = if status_text.starts_with('✓') { "#d4edda" } else { "#f8d7da" };
                                                                        rsx! {
                                                                            button {
                                                                                style: "grid-column: 1 / -1; margin-top: 6px; width: 100%; padding: 4px 8px; border: 1px solid #28a745; background: #28a745; color: white; cursor: pointer; border-radius: 3px; font-size: 11px;",
                                                                                onclick: move |_| {
                                                                                    let form_opt4 = state.get_current_form();
                                                                                    if let Some(form) = form_opt4.as_ref() {
                                                                                        if let Some(ctrl) = form.get_control(cid) {
                                                                                            let cs = ctrl.properties.get_string("ConnectionString")
                                                                                                .unwrap_or("").to_string();
                                                                                            if cs.is_empty() {
                                                                                                conn_status.set("⚠ No connection string".to_string());
                                                                                                return;
                                                                                            }
                                                                                            let cs = resolve_conn_str(&cs);
                                                                                            match vybe_runtime::data_access::test_connection_and_list_tables(&cs) {
                                                                                                Ok(tbl_list) => {
                                                                                                    conn_status.set(format!("✓ Connected — {} tables found", tbl_list.len()));
                                                                                                    table_list.set(tbl_list);
                                                                                                }
                                                                                                Err(e) => {
                                                                                                    conn_status.set(format!("✗ {}", e));
                                                                                                    table_list.set(Vec::new());
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                },
                                                                                "⚡ Test Connection & Fetch Tables"
                                                                            }

                                                                            if !status_text.is_empty() {
                                                                                div { style: "grid-column: 1 / -1; font-size: 10px; padding: 3px 6px; border-radius: 3px; margin-top: 2px; background: {status_bg};",
                                                                                    "{status_text}"
                                                                                }
                                                                            }

                                                                            if !tables.is_empty() {
                                                                                div { style: "font-weight: bold; margin-top: 4px;", "Table" }
                                                                                select {
                                                                                    style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                                                    onchange: move |evt| {
                                                                                        let tbl = evt.value();
                                                                                        if !tbl.is_empty() {
                                                                                            let select_cmd = format!("SELECT * FROM {}", tbl);
                                                                                            state.update_control_property(cid, "SelectCommand", select_cmd);
                                                                                        }
                                                                                    },
                                                                                    option { value: "", "— pick a table —" }
                                                                                    for tbl_name in &tables {
                                                                                        option { value: "{tbl_name}", "{tbl_name}" }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            // URL property for WebBrowser
                                            if matches!(control.control_type, vybe_forms::ControlType::WebBrowser) {
                                                {
                                                    let url = control.properties.get_string("URL").map(|s| s.to_string()).unwrap_or_else(|| "about:blank".to_string());
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "URL" }
                                                        input {
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{url}",
                                                            placeholder: "about:blank",
                                                            oninput: move |evt| {
                                                                state.update_control_property(cid, "URL", evt.value());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            // HTML and ToolbarVisible properties for RichTextBox
                                            if matches!(control.control_type, vybe_forms::ControlType::RichTextBox) {
                                                {
                                                    let html = control.properties.get_string("HTML").map(|s| s.to_string()).unwrap_or_else(|| "".to_string());
                                                    let toolbar_visible = control.properties.get_bool("ToolbarVisible").unwrap_or(true);
                                                    let rtb_prop_id = format!("rtb_prop_{}", cid);
                                                    let rtb_prop_id_bold = rtb_prop_id.clone();
                                                    let rtb_prop_id_italic = rtb_prop_id.clone();
                                                    let rtb_prop_id_underline = rtb_prop_id.clone();
                                                    let rtb_prop_id_ul = rtb_prop_id.clone();
                                                    let rtb_prop_id_ol = rtb_prop_id.clone();
                                                    let rtb_prop_id_mount = rtb_prop_id.clone();
                                                    let rtb_prop_id_input = rtb_prop_id.clone();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "HTML Editor" }
                                                        div {
                                                            style: "border: 1px solid #ccc; background: white; margin-bottom: 8px;",
                                                            // Toolbar
                                                            div {
                                                                style: "display: flex; gap: 2px; padding: 4px; background: #f0f0f0; border-bottom: 1px solid #ccc;",
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-weight: bold;",
                                                                    title: "Bold (Ctrl+B)",
                                                                    onclick: move |_| {
                                                                        let _ = document::eval(&format!("document.execCommand('bold', false, null); document.getElementById('{}').focus();", rtb_prop_id_bold));
                                                                    },
                                                                    "B"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-style: italic;",
                                                                    title: "Italic (Ctrl+I)",
                                                                    onclick: move |_| {
                                                                        let _ = document::eval(&format!("document.execCommand('italic', false, null); document.getElementById('{}').focus();", rtb_prop_id_italic));
                                                                    },
                                                                    "I"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; text-decoration: underline;",
                                                                    title: "Underline (Ctrl+U)",
                                                                    onclick: move |_| {
                                                                        let _ = document::eval(&format!("document.execCommand('underline', false, null); document.getElementById('{}').focus();", rtb_prop_id_underline));
                                                                    },
                                                                    "U"
                                                                }
                                                                div { style: "width: 1px; background: #ccc; margin: 0 4px;" }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                    title: "Bullet List",
                                                                    onclick: move |_| {
                                                                        let _ = document::eval(&format!("document.execCommand('insertUnorderedList', false, null); document.getElementById('{}').focus();", rtb_prop_id_ul));
                                                                    },
                                                                    "• List"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                    title: "Numbered List",
                                                                    onclick: move |_| {
                                                                        let _ = document::eval(&format!("document.execCommand('insertOrderedList', false, null); document.getElementById('{}').focus();", rtb_prop_id_ol));
                                                                    },
                                                                    "1. List"
                                                                }
                                                            }
                                                            // ContentEditable div
                                                            div {
                                                                id: "{rtb_prop_id}",
                                                                contenteditable: "true",
                                                                style: "min-height: 100px; max-height: 200px; padding: 8px; overflow: auto; outline: none;",
                                                                dangerous_inner_html: "{html}",
                                                                onmounted: move |_| {
                                                                    // Add keyboard shortcuts without format! brace escaping headaches
                                                                    let js = r#"(function() {
                                                                        const editor = document.getElementById('__ID__');
                                                                        if (editor) {
                                                                            editor.addEventListener('keydown', function(e) {
                                                                                if (e.ctrlKey || e.metaKey) {
                                                                                    switch(e.key.toLowerCase()) {
                                                                                        case 'b':
                                                                                            e.preventDefault();
                                                                                            document.execCommand('bold', false, null);
                                                                                            break;
                                                                                        case 'i':
                                                                                            e.preventDefault();
                                                                                            document.execCommand('italic', false, null);
                                                                                            break;
                                                                                        case 'u':
                                                                                            e.preventDefault();
                                                                                            document.execCommand('underline', false, null);
                                                                                            break;
                                                                                    }
                                                                                }
                                                                            });
                                                                        }
                                                                    })();"#
                                                                        .replace("__ID__", &rtb_prop_id_mount);
                                                                    let _ = document::eval(&js);
                                                                },
                                                                oninput: move |_| {
                                                                    // Update the HTML property when content changes
                                                                    let rtb_id_clone = rtb_prop_id_input.clone();
                                                                    let ctrl_id = cid;
                                                                    spawn(async move {
                                                                        let js = r#"(function() {
                                                                            const editor = document.getElementById('__ID__');
                                                                            if (editor) {
                                                                                return editor.innerHTML;
                                                                            } else {
                                                                                return '';
                                                                            }
                                                                        })()"#
                                                                            .replace("__ID__", &rtb_id_clone);

                                                                        if let Ok(result) = document::eval(&js).recv::<serde_json::Value>().await {
                                                                            if let Some(html_value) = result.as_str() {
                                                                                state.update_control_property(ctrl_id, "HTML", html_value.to_string());
                                                                            }
                                                                        }
                                                                    });
                                                                },
                                                            }
                                                        }
                                                        
                                                        div { style: "font-weight: bold; margin-top: 8px;", "ToolbarVisible" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: toolbar_visible,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "ToolbarVisible", evt.checked().to_string());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            // List property for ListBox and ComboBox
                                            if matches!(control.control_type, vybe_forms::ControlType::ListBox | vybe_forms::ControlType::ComboBox) {
                                                {
                                                    let list_items = control.get_list_items();
                                                    let list_text = list_items.join("\n");
                                                    
                                                    rsx! {
                                                        div { style: "font-weight: bold; grid-column: 1 / -1; margin-top: 8px;", "List Items (one per line)" }
                                                        textarea { 
                                                            style: "grid-column: 1 / -1; width: 100%; height: 100px; border: 1px solid #ccc; padding: 4px; font-size: 12px; font-family: monospace; resize: vertical;",
                                                            initial_value: "{list_text}",
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "List", evt.value());
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            // ── Control-Specific Properties ──
                                            // CheckBox / RadioButton: Checked
                                            if matches!(control.control_type, vybe_forms::ControlType::CheckBox | vybe_forms::ControlType::RadioButton) {
                                                {
                                                    let is_checked = control.properties.get_bool("Checked").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Checked" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: is_checked,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "Checked", evt.checked().to_string());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // CheckBox: CheckState (Unchecked=0, Checked=1, Indeterminate=2)
                                            if matches!(control.control_type, vybe_forms::ControlType::CheckBox) {
                                                {
                                                    let check_state = control.properties.get_int("CheckState").unwrap_or(
                                                        if control.properties.get_bool("Checked").unwrap_or(false) { 1 } else { 0 }
                                                    );
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "CheckState" }
                                                        select {
                                                            value: "{check_state}",
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "CheckState", evt.value());
                                                            },
                                                            option { value: "0", "Unchecked" }
                                                            option { value: "1", "Checked" }
                                                            option { value: "2", "Indeterminate" }
                                                        }
                                                    }
                                                }
                                            }
                                            // CheckBox: ThreeState
                                            if matches!(control.control_type, vybe_forms::ControlType::CheckBox) {
                                                {
                                                    let three_state = control.properties.get_bool("ThreeState").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "ThreeState" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: three_state,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "ThreeState", evt.checked().to_string());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // TextBox: Multiline, ReadOnly, MaxLength, PasswordChar
                                            if matches!(control.control_type, vybe_forms::ControlType::TextBox) {
                                                {
                                                    let multiline = control.properties.get_bool("Multiline").unwrap_or(false);
                                                    let readonly = control.properties.get_bool("ReadOnly").unwrap_or(false);
                                                    let max_length = control.properties.get_int("MaxLength").unwrap_or(32767);
                                                    let password_char = control.properties.get_string("PasswordChar").map(|s| s.to_string()).unwrap_or_default();
                                                    let max_length_str = max_length.to_string();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Multiline" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: multiline,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "Multiline", evt.checked().to_string());
                                                            }
                                                        }
                                                        div { style: "font-weight: bold;", "ReadOnly" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: readonly,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "ReadOnly", evt.checked().to_string());
                                                            }
                                                        }
                                                        div { style: "font-weight: bold;", "MaxLength" }
                                                        input {
                                                            r#type: "number",
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{max_length_str}",
                                                            oninput: move |evt| {
                                                                state.update_control_property(cid, "MaxLength", evt.value());
                                                            }
                                                        }
                                                        div { style: "font-weight: bold;", "PasswordChar" }
                                                        input {
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{password_char}",
                                                            maxlength: "1",
                                                            oninput: move |evt| {
                                                                state.update_control_property(cid, "PasswordChar", evt.value());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // ComboBox: DropDownStyle, Sorted
                                            if matches!(control.control_type, vybe_forms::ControlType::ComboBox) {
                                                {
                                                    let ddstyle = control.properties.get_int("DropDownStyle").unwrap_or(0);
                                                    let sorted = control.properties.get_bool("Sorted").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "DropDownStyle" }
                                                        select {
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{ddstyle}",
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "DropDownStyle", evt.value());
                                                            },
                                                            option { value: "0", "Simple" }
                                                            option { value: "1", "DropDown" }
                                                            option { value: "2", "DropDownList" }
                                                        }
                                                        div { style: "font-weight: bold;", "Sorted" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: sorted,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "Sorted", evt.checked().to_string());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // ListBox: SelectionMode, Sorted
                                            if matches!(control.control_type, vybe_forms::ControlType::ListBox) {
                                                {
                                                    let sel_mode = control.properties.get_int("SelectionMode").unwrap_or(1);
                                                    let sorted = control.properties.get_bool("Sorted").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "SelectionMode" }
                                                        select {
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{sel_mode}",
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "SelectionMode", evt.value());
                                                            },
                                                            option { value: "0", "None" }
                                                            option { value: "1", "One" }
                                                            option { value: "2", "MultiSimple" }
                                                            option { value: "3", "MultiExtended" }
                                                        }
                                                        div { style: "font-weight: bold;", "Sorted" }
                                                        input {
                                                            r#type: "checkbox",
                                                            checked: sorted,
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "Sorted", evt.checked().to_string());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            // ProgressBar: Value, Minimum, Maximum, Step
                                            if matches!(control.control_type, vybe_forms::ControlType::ProgressBar) {
                                                {
                                                    let val = control.properties.get_int("Value").unwrap_or(0).to_string();
                                                    let min = control.properties.get_int("Minimum").unwrap_or(0).to_string();
                                                    let max = control.properties.get_int("Maximum").unwrap_or(100).to_string();
                                                    let step = control.properties.get_int("Step").unwrap_or(10).to_string();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Value" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{val}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Value", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Minimum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{min}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Minimum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Maximum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{max}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Maximum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Step" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{step}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Step", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // NumericUpDown: Value, Minimum, Maximum, Increment, DecimalPlaces
                                            if matches!(control.control_type, vybe_forms::ControlType::NumericUpDown) {
                                                {
                                                    let val = control.properties.get_int("Value").unwrap_or(0).to_string();
                                                    let min = control.properties.get_int("Minimum").unwrap_or(0).to_string();
                                                    let max = control.properties.get_int("Maximum").unwrap_or(100).to_string();
                                                    let inc = control.properties.get_int("Increment").unwrap_or(1).to_string();
                                                    let dec = control.properties.get_int("DecimalPlaces").unwrap_or(0).to_string();
                                                    let readonly = control.properties.get_bool("ReadOnly").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Value" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{val}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Value", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Minimum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{min}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Minimum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Maximum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{max}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Maximum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Increment" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{inc}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Increment", evt.value()); } }
                                                        div { style: "font-weight: bold;", "DecimalPlaces" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{dec}",
                                                            oninput: move |evt| { state.update_control_property(cid, "DecimalPlaces", evt.value()); } }
                                                        div { style: "font-weight: bold;", "ReadOnly" }
                                                        input { r#type: "checkbox", checked: readonly,
                                                            onchange: move |evt| { state.update_control_property(cid, "ReadOnly", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // TrackBar: Value, Minimum, Maximum, TickFrequency, SmallChange, LargeChange
                                            if matches!(control.control_type, vybe_forms::ControlType::TrackBar) {
                                                {
                                                    let val = control.properties.get_int("Value").unwrap_or(0).to_string();
                                                    let min = control.properties.get_int("Minimum").unwrap_or(0).to_string();
                                                    let max = control.properties.get_int("Maximum").unwrap_or(10).to_string();
                                                    let tick = control.properties.get_int("TickFrequency").unwrap_or(1).to_string();
                                                    let sm = control.properties.get_int("SmallChange").unwrap_or(1).to_string();
                                                    let lg = control.properties.get_int("LargeChange").unwrap_or(5).to_string();
                                                    let orient = control.properties.get_string("Orientation").map(|s| s.to_string()).unwrap_or_else(|| "Horizontal".to_string());
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Value" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{val}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Value", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Minimum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{min}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Minimum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Maximum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{max}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Maximum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "TickFrequency" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{tick}",
                                                            oninput: move |evt| { state.update_control_property(cid, "TickFrequency", evt.value()); } }
                                                        div { style: "font-weight: bold;", "SmallChange" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{sm}",
                                                            oninput: move |evt| { state.update_control_property(cid, "SmallChange", evt.value()); } }
                                                        div { style: "font-weight: bold;", "LargeChange" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{lg}",
                                                            oninput: move |evt| { state.update_control_property(cid, "LargeChange", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Orientation" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{orient}",
                                                            onchange: move |evt| { state.update_control_property(cid, "Orientation", evt.value()); },
                                                            option { value: "Horizontal", "Horizontal" }
                                                            option { value: "Vertical", "Vertical" }
                                                        }
                                                    }
                                                }
                                            }
                                            // HScrollBar / VScrollBar: Value, Minimum, Maximum, SmallChange, LargeChange
                                            if matches!(control.control_type, vybe_forms::ControlType::HScrollBar | vybe_forms::ControlType::VScrollBar) {
                                                {
                                                    let val = control.properties.get_int("Value").unwrap_or(0).to_string();
                                                    let min = control.properties.get_int("Minimum").unwrap_or(0).to_string();
                                                    let max = control.properties.get_int("Maximum").unwrap_or(100).to_string();
                                                    let sm = control.properties.get_int("SmallChange").unwrap_or(1).to_string();
                                                    let lg = control.properties.get_int("LargeChange").unwrap_or(10).to_string();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Value" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{val}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Value", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Minimum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{min}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Minimum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "Maximum" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{max}",
                                                            oninput: move |evt| { state.update_control_property(cid, "Maximum", evt.value()); } }
                                                        div { style: "font-weight: bold;", "SmallChange" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{sm}",
                                                            oninput: move |evt| { state.update_control_property(cid, "SmallChange", evt.value()); } }
                                                        div { style: "font-weight: bold;", "LargeChange" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{lg}",
                                                            oninput: move |evt| { state.update_control_property(cid, "LargeChange", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // DateTimePicker: Format, ShowCheckBox, ShowUpDown, CustomFormat
                                            if matches!(control.control_type, vybe_forms::ControlType::DateTimePicker) {
                                                {
                                                    let fmt = control.properties.get_string("Format").map(|s| s.to_string()).unwrap_or_else(|| "Long".to_string());
                                                    let custom_fmt = control.properties.get_string("CustomFormat").map(|s| s.to_string()).unwrap_or_default();
                                                    let show_cb = control.properties.get_bool("ShowCheckBox").unwrap_or(false);
                                                    let show_ud = control.properties.get_bool("ShowUpDown").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Format" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{fmt}",
                                                            onchange: move |evt| { state.update_control_property(cid, "Format", evt.value()); },
                                                            option { value: "Long", "Long" }
                                                            option { value: "Short", "Short" }
                                                            option { value: "Time", "Time" }
                                                            option { value: "Custom", "Custom" }
                                                        }
                                                        div { style: "font-weight: bold;", "CustomFormat" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{custom_fmt}",
                                                            placeholder: "yyyy-MM-dd",
                                                            oninput: move |evt| { state.update_control_property(cid, "CustomFormat", evt.value()); } }
                                                        div { style: "font-weight: bold;", "ShowCheckBox" }
                                                        input { r#type: "checkbox", checked: show_cb,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowCheckBox", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShowUpDown" }
                                                        input { r#type: "checkbox", checked: show_ud,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowUpDown", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // MaskedTextBox: Mask, PromptChar
                                            if matches!(control.control_type, vybe_forms::ControlType::MaskedTextBox) {
                                                {
                                                    let mask = control.properties.get_string("Mask").map(|s| s.to_string()).unwrap_or_default();
                                                    let prompt = control.properties.get_string("PromptChar").map(|s| s.to_string()).unwrap_or_else(|| "_".to_string());
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Mask" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{mask}",
                                                            placeholder: "000-00-0000",
                                                            oninput: move |evt| { state.update_control_property(cid, "Mask", evt.value()); } }
                                                        div { style: "font-weight: bold;", "PromptChar" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{prompt}",
                                                            maxlength: "1",
                                                            oninput: move |evt| { state.update_control_property(cid, "PromptChar", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // SplitContainer: Orientation, SplitterDistance, FixedPanel, IsSplitterFixed
                                            if matches!(control.control_type, vybe_forms::ControlType::SplitContainer) {
                                                {
                                                    let orient = control.properties.get_string("Orientation").map(|s| s.to_string()).unwrap_or_else(|| "Vertical".to_string());
                                                    let dist = control.properties.get_int("SplitterDistance").unwrap_or(100).to_string();
                                                    let fixed = control.properties.get_string("FixedPanel").map(|s| s.to_string()).unwrap_or_else(|| "None".to_string());
                                                    let is_fixed = control.properties.get_bool("IsSplitterFixed").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Orientation" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{orient}",
                                                            onchange: move |evt| { state.update_control_property(cid, "Orientation", evt.value()); },
                                                            option { value: "Vertical", "Vertical" }
                                                            option { value: "Horizontal", "Horizontal" }
                                                        }
                                                        div { style: "font-weight: bold;", "SplitterDist" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{dist}",
                                                            oninput: move |evt| { state.update_control_property(cid, "SplitterDistance", evt.value()); } }
                                                        div { style: "font-weight: bold;", "FixedPanel" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{fixed}",
                                                            onchange: move |evt| { state.update_control_property(cid, "FixedPanel", evt.value()); },
                                                            option { value: "None", "None" }
                                                            option { value: "Panel1", "Panel1" }
                                                            option { value: "Panel2", "Panel2" }
                                                        }
                                                        div { style: "font-weight: bold;", "SplitterFixed" }
                                                        input { r#type: "checkbox", checked: is_fixed,
                                                            onchange: move |evt| { state.update_control_property(cid, "IsSplitterFixed", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // FlowLayoutPanel: FlowDirection, WrapContents
                                            if matches!(control.control_type, vybe_forms::ControlType::FlowLayoutPanel) {
                                                {
                                                    let flow_dir = control.properties.get_string("FlowDirection").map(|s| s.to_string()).unwrap_or_else(|| "LeftToRight".to_string());
                                                    let wrap = control.properties.get_bool("WrapContents").unwrap_or(true);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "FlowDirection" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{flow_dir}",
                                                            onchange: move |evt| { state.update_control_property(cid, "FlowDirection", evt.value()); },
                                                            option { value: "LeftToRight", "LeftToRight" }
                                                            option { value: "TopDown", "TopDown" }
                                                            option { value: "RightToLeft", "RightToLeft" }
                                                            option { value: "BottomUp", "BottomUp" }
                                                        }
                                                        div { style: "font-weight: bold;", "WrapContents" }
                                                        input { r#type: "checkbox", checked: wrap,
                                                            onchange: move |evt| { state.update_control_property(cid, "WrapContents", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // TableLayoutPanel: ColumnCount, RowCount
                                            if matches!(control.control_type, vybe_forms::ControlType::TableLayoutPanel) {
                                                {
                                                    let cols = control.properties.get_int("ColumnCount").unwrap_or(2).to_string();
                                                    let rows = control.properties.get_int("RowCount").unwrap_or(2).to_string();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "ColumnCount" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{cols}",
                                                            oninput: move |evt| { state.update_control_property(cid, "ColumnCount", evt.value()); } }
                                                        div { style: "font-weight: bold;", "RowCount" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{rows}",
                                                            oninput: move |evt| { state.update_control_property(cid, "RowCount", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // MonthCalendar: ShowToday, ShowWeekNumbers, MaxSelectionCount
                                            if matches!(control.control_type, vybe_forms::ControlType::MonthCalendar) {
                                                {
                                                    let show_today = control.properties.get_bool("ShowToday").unwrap_or(true);
                                                    let show_weeks = control.properties.get_bool("ShowWeekNumbers").unwrap_or(false);
                                                    let max_sel = control.properties.get_int("MaxSelectionCount").unwrap_or(7).to_string();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "ShowToday" }
                                                        input { r#type: "checkbox", checked: show_today,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowToday", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShowWeekNums" }
                                                        input { r#type: "checkbox", checked: show_weeks,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowWeekNumbers", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "MaxSelection" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{max_sel}",
                                                            oninput: move |evt| { state.update_control_property(cid, "MaxSelectionCount", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // LinkLabel: LinkColor, VisitedLinkColor
                                            if matches!(control.control_type, vybe_forms::ControlType::LinkLabel) {
                                                {
                                                    let link_color = control.properties.get_string("LinkColor").map(|s| s.to_string()).unwrap_or_else(|| "#0066cc".to_string());
                                                    let visited = control.properties.get_string("VisitedLinkColor").map(|s| s.to_string()).unwrap_or_else(|| "#800080".to_string());
                                                    let link_visited = control.properties.get_bool("LinkVisited").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "LinkColor" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{link_color}",
                                                            oninput: move |evt| { state.update_control_property(cid, "LinkColor", evt.value()); } }
                                                        div { style: "font-weight: bold;", "VisitedColor" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{visited}",
                                                            oninput: move |evt| { state.update_control_property(cid, "VisitedLinkColor", evt.value()); } }
                                                        div { style: "font-weight: bold;", "LinkVisited" }
                                                        input { r#type: "checkbox", checked: link_visited,
                                                            onchange: move |evt| { state.update_control_property(cid, "LinkVisited", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // PictureBox: SizeMode, BorderStyle
                                            if matches!(control.control_type, vybe_forms::ControlType::PictureBox) {
                                                {
                                                    let mode = control.properties.get_string("SizeMode").map(|s| s.to_string()).unwrap_or_else(|| "Normal".to_string());
                                                    let bs = control.properties.get_string("BorderStyle").map(|s| s.to_string()).unwrap_or_else(|| "None".to_string());
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "SizeMode" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{mode}",
                                                            onchange: move |evt| { state.update_control_property(cid, "SizeMode", evt.value()); },
                                                            option { value: "Normal", "Normal" }
                                                            option { value: "StretchImage", "StretchImage" }
                                                            option { value: "AutoSize", "AutoSize" }
                                                            option { value: "CenterImage", "CenterImage" }
                                                            option { value: "Zoom", "Zoom" }
                                                        }
                                                        div { style: "font-weight: bold;", "BorderStyle" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{bs}",
                                                            onchange: move |evt| { state.update_control_property(cid, "BorderStyle", evt.value()); },
                                                            option { value: "None", "None" }
                                                            option { value: "FixedSingle", "FixedSingle" }
                                                            option { value: "Fixed3D", "Fixed3D" }
                                                        }
                                                    }
                                                }
                                            }
                                            // Panel: BorderStyle, AutoScroll
                                            if matches!(control.control_type, vybe_forms::ControlType::Panel) {
                                                {
                                                    let bs = control.properties.get_string("BorderStyle").map(|s| s.to_string()).unwrap_or_else(|| "None".to_string());
                                                    let autoscroll = control.properties.get_bool("AutoScroll").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "BorderStyle" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{bs}",
                                                            onchange: move |evt| { state.update_control_property(cid, "BorderStyle", evt.value()); },
                                                            option { value: "None", "None" }
                                                            option { value: "FixedSingle", "FixedSingle" }
                                                            option { value: "Fixed3D", "Fixed3D" }
                                                        }
                                                        div { style: "font-weight: bold;", "AutoScroll" }
                                                        input { r#type: "checkbox", checked: autoscroll,
                                                            onchange: move |evt| { state.update_control_property(cid, "AutoScroll", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // ToolStripMenuItem: Checked, CheckOnClick, ShortcutKeys
                                            if matches!(control.control_type, vybe_forms::ControlType::ToolStripMenuItem) {
                                                {
                                                    let checked = control.properties.get_bool("Checked").unwrap_or(false);
                                                    let check_on_click = control.properties.get_bool("CheckOnClick").unwrap_or(false);
                                                    let shortcut = control.properties.get_string("ShortcutKeys").map(|s| s.to_string()).unwrap_or_default();
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Checked" }
                                                        input { r#type: "checkbox", checked: checked,
                                                            onchange: move |evt| { state.update_control_property(cid, "Checked", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "CheckOnClick" }
                                                        input { r#type: "checkbox", checked: check_on_click,
                                                            onchange: move |evt| { state.update_control_property(cid, "CheckOnClick", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShortcutKeys" }
                                                        input { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{shortcut}",
                                                            placeholder: "Ctrl+S",
                                                            oninput: move |evt| { state.update_control_property(cid, "ShortcutKeys", evt.value()); } }
                                                    }
                                                }
                                            }
                                            // ToolTip: AutoPopDelay, InitialDelay, ShowAlways
                                            if matches!(control.control_type, vybe_forms::ControlType::ToolTip) {
                                                {
                                                    let autopop = control.properties.get_int("AutoPopDelay").unwrap_or(5000).to_string();
                                                    let initial = control.properties.get_int("InitialDelay").unwrap_or(500).to_string();
                                                    let show_always = control.properties.get_bool("ShowAlways").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "AutoPopDelay" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{autopop}",
                                                            oninput: move |evt| { state.update_control_property(cid, "AutoPopDelay", evt.value()); } }
                                                        div { style: "font-weight: bold;", "InitialDelay" }
                                                        input { r#type: "number", style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{initial}",
                                                            oninput: move |evt| { state.update_control_property(cid, "InitialDelay", evt.value()); } }
                                                        div { style: "font-weight: bold;", "ShowAlways" }
                                                        input { r#type: "checkbox", checked: show_always,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowAlways", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // TreeView: CheckBoxes, ShowLines, ShowRootLines, ShowPlusMinus, LabelEdit
                                            if matches!(control.control_type, vybe_forms::ControlType::TreeView) {
                                                {
                                                    let cbs = control.properties.get_bool("CheckBoxes").unwrap_or(false);
                                                    let lines = control.properties.get_bool("ShowLines").unwrap_or(true);
                                                    let root_lines = control.properties.get_bool("ShowRootLines").unwrap_or(true);
                                                    let plus_minus = control.properties.get_bool("ShowPlusMinus").unwrap_or(true);
                                                    let label_edit = control.properties.get_bool("LabelEdit").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "CheckBoxes" }
                                                        input { r#type: "checkbox", checked: cbs,
                                                            onchange: move |evt| { state.update_control_property(cid, "CheckBoxes", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShowLines" }
                                                        input { r#type: "checkbox", checked: lines,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowLines", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShowRootLines" }
                                                        input { r#type: "checkbox", checked: root_lines,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowRootLines", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "ShowPlusMinus" }
                                                        input { r#type: "checkbox", checked: plus_minus,
                                                            onchange: move |evt| { state.update_control_property(cid, "ShowPlusMinus", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "LabelEdit" }
                                                        input { r#type: "checkbox", checked: label_edit,
                                                            onchange: move |evt| { state.update_control_property(cid, "LabelEdit", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // ListView: View, FullRowSelect, GridLines, CheckBoxes, MultiSelect, LabelEdit
                                            if matches!(control.control_type, vybe_forms::ControlType::ListView) {
                                                {
                                                    let view = control.properties.get_int("View").unwrap_or(1);
                                                    let full_row = control.properties.get_bool("FullRowSelect").unwrap_or(false);
                                                    let grid_lines = control.properties.get_bool("GridLines").unwrap_or(false);
                                                    let cbs = control.properties.get_bool("CheckBoxes").unwrap_or(false);
                                                    let multi = control.properties.get_bool("MultiSelect").unwrap_or(true);
                                                    let label_edit = control.properties.get_bool("LabelEdit").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "View" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{view}",
                                                            onchange: move |evt| { state.update_control_property(cid, "View", evt.value()); },
                                                            option { value: "0", "LargeIcon" }
                                                            option { value: "1", "Details" }
                                                            option { value: "2", "SmallIcon" }
                                                            option { value: "3", "List" }
                                                            option { value: "4", "Tile" }
                                                        }
                                                        div { style: "font-weight: bold;", "FullRowSelect" }
                                                        input { r#type: "checkbox", checked: full_row,
                                                            onchange: move |evt| { state.update_control_property(cid, "FullRowSelect", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "GridLines" }
                                                        input { r#type: "checkbox", checked: grid_lines,
                                                            onchange: move |evt| { state.update_control_property(cid, "GridLines", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "CheckBoxes" }
                                                        input { r#type: "checkbox", checked: cbs,
                                                            onchange: move |evt| { state.update_control_property(cid, "CheckBoxes", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "MultiSelect" }
                                                        input { r#type: "checkbox", checked: multi,
                                                            onchange: move |evt| { state.update_control_property(cid, "MultiSelect", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "LabelEdit" }
                                                        input { r#type: "checkbox", checked: label_edit,
                                                            onchange: move |evt| { state.update_control_property(cid, "LabelEdit", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // DataGridView: ReadOnly, AllowUserToAddRows/DeleteRows, AutoGenerateColumns
                                            if matches!(control.control_type, vybe_forms::ControlType::DataGridView) {
                                                {
                                                    let readonly = control.properties.get_bool("ReadOnly").unwrap_or(false);
                                                    let add_rows = control.properties.get_bool("AllowUserToAddRows").unwrap_or(true);
                                                    let del_rows = control.properties.get_bool("AllowUserToDeleteRows").unwrap_or(true);
                                                    let auto_gen = control.properties.get_bool("AutoGenerateColumns").unwrap_or(true);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "ReadOnly" }
                                                        input { r#type: "checkbox", checked: readonly,
                                                            onchange: move |evt| { state.update_control_property(cid, "ReadOnly", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "AllowAddRows" }
                                                        input { r#type: "checkbox", checked: add_rows,
                                                            onchange: move |evt| { state.update_control_property(cid, "AllowUserToAddRows", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "AllowDelRows" }
                                                        input { r#type: "checkbox", checked: del_rows,
                                                            onchange: move |evt| { state.update_control_property(cid, "AllowUserToDeleteRows", evt.checked().to_string()); } }
                                                        div { style: "font-weight: bold;", "AutoGenCols" }
                                                        input { r#type: "checkbox", checked: auto_gen,
                                                            onchange: move |evt| { state.update_control_property(cid, "AutoGenerateColumns", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                            // TabControl: Alignment, Multiline
                                            if matches!(control.control_type, vybe_forms::ControlType::TabControl) {
                                                {
                                                    let alignment = control.properties.get_string("Alignment").map(|s| s.to_string()).unwrap_or_else(|| "Top".to_string());
                                                    let multiline = control.properties.get_bool("Multiline").unwrap_or(false);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Alignment" }
                                                        select { style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;", value: "{alignment}",
                                                            onchange: move |evt| { state.update_control_property(cid, "Alignment", evt.value()); },
                                                            option { value: "Top", "Top" }
                                                            option { value: "Bottom", "Bottom" }
                                                            option { value: "Left", "Left" }
                                                            option { value: "Right", "Right" }
                                                        }
                                                        div { style: "font-weight: bold;", "Multiline" }
                                                        input { r#type: "checkbox", checked: multiline,
                                                            onchange: move |evt| { state.update_control_property(cid, "Multiline", evt.checked().to_string()); } }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                "Control not found"
                            }
                        } else {
                            "No form loaded"
                        }
                    } else {
                        if let Some(form) = form_opt.as_ref() {
                            {
                                let form_caption = form.text.clone();
                                let form_width = form.width;
                                let form_height = form.height;
                                let back_color = form.back_color.clone().unwrap_or_else(|| "#f8fafc".to_string());
                                let fore_color = form.fore_color.clone().unwrap_or_else(|| "#0f172a".to_string());
                                let font = form.font.clone().unwrap_or_else(|| "Segoe UI, 12px".to_string());

                                // Parse font size
                                let mut font_parts = font.split(',').map(|s| s.trim());
                                let font_family = font_parts.next().unwrap_or("Segoe UI").to_string();
                                let font_size_part = font_parts.next().unwrap_or("12px");
                                let font_size_num: String = font_size_part.trim_end_matches("px").trim_end_matches("pt").to_string();
                                let _font_family_sel = font_family.clone();
                                let font_family_sel2 = font_family.clone();
                                let font_size_sel = font_size_num.clone();
                                let _font_size_sel2 = font_size_num.clone();

                                rsx! {
                                    div { style: "display: grid; grid-template-columns: 90px 1fr; gap: 4px; align-items: center;",
                                        div { style: "font-weight: bold;", "Form" }
                                        div { style: "font-size: 12px; color: #555;", "{form_caption}" }

                                        div { style: "font-weight: bold;", "Text" }
                                        input {
                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                            value: "{form_caption}",
                                            oninput: move |evt| {
                                                state.update_form_property("Text", evt.value());
                                            }
                                        }

                                        div { style: "font-weight: bold;", "Width" }
                                        input {
                                            r#type: "number",
                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                            value: "{form_width}",
                                            oninput: move |evt| {
                                                state.update_form_property("Width", evt.value());
                                            }
                                        }

                                        div { style: "font-weight: bold;", "Height" }
                                        input {
                                            r#type: "number",
                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                            value: "{form_height}",
                                            oninput: move |evt| {
                                                state.update_form_property("Height", evt.value());
                                            }
                                        }

                                        div { style: "font-weight: bold;", "BackColor" }
                                        div { style: "display: flex; align-items: center; gap: 8px;",
                                            input {
                                                r#type: "color",
                                                value: if back_color.starts_with('#') && back_color.len() == 7 { back_color.clone() } else { "#f8fafc".to_string() },
                                                style: "width: 46px; height: 28px; padding: 0; border: 1px solid #ccc; background: transparent;",
                                                onchange: move |evt| {
                                                    state.update_form_property("BackColor", evt.value());
                                                }
                                            }
                                            input {
                                                style: "flex: 1; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{back_color}",
                                                placeholder: "#RRGGBB or css color",
                                                oninput: move |evt| {
                                                    state.update_form_property("BackColor", evt.value());
                                                }
                                            }
                                        }

                                        div { style: "font-weight: bold; margin-top: 4px;", "ForeColor" }
                                        div { style: "display: flex; align-items: center; gap: 8px;",
                                            input {
                                                r#type: "color",
                                                value: if fore_color.starts_with('#') && fore_color.len() == 7 { fore_color.clone() } else { "#0f172a".to_string() },
                                                style: "width: 46px; height: 28px; padding: 0; border: 1px solid #ccc; background: transparent;",
                                                onchange: move |evt| {
                                                    state.update_form_property("ForeColor", evt.value());
                                                }
                                            }
                                            input {
                                                style: "flex: 1; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                value: "{fore_color}",
                                                placeholder: "#RRGGBB or css color",
                                                oninput: move |evt| {
                                                    state.update_form_property("ForeColor", evt.value());
                                                }
                                            }
                                        }

                                        div { style: "font-weight: bold; margin-top: 4px;", "Font" }
                                        div { style: "display: flex; gap: 6px; align-items: center;",
                                            select {
                                                style: "flex: 1; border: 1px solid #ccc; padding: 4px; font-size: 12px;",
                                                value: "{font_family}",
                                                onchange: move |evt| {
                                                    let fam = evt.value();
                                                    let size = font_size_sel.clone();
                                                    state.update_form_property("Font", format!("{}, {}px", fam, size));
                                                },
                                                option { value: "Segoe UI", "Segoe UI" }
                                                option { value: "Arial", "Arial" }
                                                option { value: "Helvetica", "Helvetica" }
                                                option { value: "Times New Roman", "Times New Roman" }
                                                option { value: "Courier New", "Courier New" }
                                                option { value: "Consolas", "Consolas" }
                                                option { value: "Menlo", "Menlo" }
                                                option { value: "Monaco", "Monaco" }
                                                option { value: "Inter", "Inter" }
                                                option { value: "Roboto", "Roboto" }
                                            }
                                            select {
                                                style: "width: 90px; border: 1px solid #ccc; padding: 4px; font-size: 12px;",
                                                value: "{font_size_num}",
                                                onchange: move |evt| {
                                                    let size = evt.value();
                                                    let fam = font_family_sel2.clone();
                                                    state.update_form_property("Font", format!("{}, {}px", fam, size));
                                                },
                                                option { value: "10", "10" }
                                                option { value: "11", "11" }
                                                option { value: "12", "12" }
                                                option { value: "14", "14" }
                                                option { value: "16", "16" }
                                                option { value: "18", "18" }
                                                option { value: "20", "20" }
                                            }
                                        }
                                    }
                                }
                            }
                        } else {
                            "No Selection"
                        }
                    }
                }
            }
        }
    }
}
