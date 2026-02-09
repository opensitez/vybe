use dioxus::prelude::*;
use crate::app_state::AppState;
use irys_parser::{parse_program, Declaration};

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
                if let irys_parser::ast::decl::MethodDecl::Sub(sub) = method {
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
    let mut state = use_context::<AppState>();
    let selected_control_id = *state.selected_control.read();
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
                                                irys_forms::ControlType::Button => vec!["Click", "DblClick", "MouseDown", "MouseUp", "MouseMove"],
                                                irys_forms::ControlType::TextBox => vec!["Change", "KeyPress", "GotFocus", "LostFocus"],
                                                irys_forms::ControlType::Label => vec!["Click", "DblClick"],
                                                irys_forms::ControlType::CheckBox => vec!["Click"],
                                                irys_forms::ControlType::RadioButton => vec!["Click"],
                                                irys_forms::ControlType::ListBox => vec!["Click", "DblClick"],
                                                irys_forms::ControlType::ComboBox => vec!["Click", "Change", "DropDown"],
                                                irys_forms::ControlType::Frame => vec!["Click", "DblClick"],
                                                irys_forms::ControlType::TreeView => vec!["Click", "DblClick"],
                                                irys_forms::ControlType::DataGridView => vec!["Click"],
                                                irys_forms::ControlType::Panel => vec!["Click", "DblClick"],
                                                irys_forms::ControlType::ListView => vec!["Click", "DblClick"],
                                                _ => vec!["Click"],
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
                                                                    let params = if is_arr { "Index As Integer" } else { "" };
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
                                                                            let _ = eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
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
                                                                        let _ = eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
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
                            let events = vec!["Load", "Click"]; // basic form events
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
                                                    let params = "";
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
                                                            let _ = eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
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
                                                        let _ = eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
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
                                    
                                    let caption = control.get_caption().map(|s| s.to_string()).unwrap_or_default();
                                    let text = control.get_text().map(|s| s.to_string()).unwrap_or_default();
                                    let back_color = control.get_back_color().map(|s| s.to_string()).unwrap_or_else(|| "#f8fafc".to_string());
                                    let fore_color = control.get_fore_color().map(|s| s.to_string()).unwrap_or_else(|| "#0f172a".to_string());
                                    let font = control.get_font().map(|s| s.to_string()).unwrap_or_else(|| "Segoe UI, 12px".to_string());

                                    // Parse font into family + size
                                    let mut font_parts = font.split(',').map(|s| s.trim());
                                    let font_family = font_parts.next().unwrap_or("Segoe UI").to_string();
                                    let font_size_part = font_parts.next().unwrap_or("12px");
                                    let font_size_num: String = font_size_part.trim_end_matches("px").trim_end_matches("pt").to_string();
                                    let font_family_sel = font_family.clone();
                                    let font_family_sel2 = font_family.clone();
                                    let font_size_sel = font_size_num.clone();
                                    let font_size_sel2 = font_size_num.clone();
                                    
                                    let index_str = control.index.map(|i| i.to_string()).unwrap_or_default();
                                    let has_caption = matches!(control.control_type,
                                        irys_forms::ControlType::Button |
                                        irys_forms::ControlType::Label |
                                        irys_forms::ControlType::CheckBox |
                                        irys_forms::ControlType::Frame);
                                    let has_text = matches!(control.control_type, irys_forms::ControlType::TextBox);

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
                                                div { style: "font-weight: bold;", "Caption" }
                                                input { 
                                                    style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                    value: "{caption}",
                                                    oninput: move |evt| {
                                                        state.update_control_property(cid, "Caption", evt.value());
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
                                            
                                            // CheckBox Value property
                                            if matches!(control.control_type, irys_forms::ControlType::CheckBox | irys_forms::ControlType::RadioButton) {
                                                {
                                                    let value = control.properties.get_int("Value").unwrap_or(0);
                                                    rsx! {
                                                        div { style: "font-weight: bold;", "Value" }
                                                        select {
                                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                                            value: "{value}",
                                                            onchange: move |evt| {
                                                                state.update_control_property(cid, "Value", evt.value());
                                                            },
                                                            option { value: "0", "Unchecked" }
                                                            option { value: "1", "Checked" }
                                                        }
                                                    }
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
                                            
                                            // URL property for WebBrowser
                                            if matches!(control.control_type, irys_forms::ControlType::WebBrowser) {
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
                                            if matches!(control.control_type, irys_forms::ControlType::RichTextBox) {
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
                                                                        let _ = eval(&format!("document.execCommand('bold', false, null); document.getElementById('{}').focus();", rtb_prop_id_bold));
                                                                    },
                                                                    "B"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-style: italic;",
                                                                    title: "Italic (Ctrl+I)",
                                                                    onclick: move |_| {
                                                                        let _ = eval(&format!("document.execCommand('italic', false, null); document.getElementById('{}').focus();", rtb_prop_id_italic));
                                                                    },
                                                                    "I"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; text-decoration: underline;",
                                                                    title: "Underline (Ctrl+U)",
                                                                    onclick: move |_| {
                                                                        let _ = eval(&format!("document.execCommand('underline', false, null); document.getElementById('{}').focus();", rtb_prop_id_underline));
                                                                    },
                                                                    "U"
                                                                }
                                                                div { style: "width: 1px; background: #ccc; margin: 0 4px;" }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                    title: "Bullet List",
                                                                    onclick: move |_| {
                                                                        let _ = eval(&format!("document.execCommand('insertUnorderedList', false, null); document.getElementById('{}').focus();", rtb_prop_id_ul));
                                                                    },
                                                                    "• List"
                                                                }
                                                                button {
                                                                    style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                    title: "Numbered List",
                                                                    onclick: move |_| {
                                                                        let _ = eval(&format!("document.execCommand('insertOrderedList', false, null); document.getElementById('{}').focus();", rtb_prop_id_ol));
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
                                                                    let _ = eval(&js);
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

                                                                        if let Ok(result) = eval(&js).recv().await {
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
                                            if matches!(control.control_type, irys_forms::ControlType::ListBox | irys_forms::ControlType::ComboBox) {
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
                                let form_caption = form.caption.clone();
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
                                let font_family_sel = font_family.clone();
                                let font_family_sel2 = font_family.clone();
                                let font_size_sel = font_size_num.clone();
                                let font_size_sel2 = font_size_num.clone();

                                rsx! {
                                    div { style: "display: grid; grid-template-columns: 90px 1fr; gap: 4px; align-items: center;",
                                        div { style: "font-weight: bold;", "Form" }
                                        div { style: "font-size: 12px; color: #555;", "{form_caption}" }

                                        div { style: "font-weight: bold;", "Caption" }
                                        input {
                                            style: "width: 100%; border: 1px solid #ccc; padding: 2px 4px; font-size: 12px;",
                                            value: "{form_caption}",
                                            oninput: move |evt| {
                                                state.update_form_property("Caption", evt.value());
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
