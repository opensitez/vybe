use dioxus::prelude::*;
use crate::app_state::AppState;
use irys_forms::{Form, EventType};
                            // Controls
                            for control in form.controls {
                                { render_control_with_handler(&control, on_event.clone()) }
                            }
                                background: {form_back};
                                color: {form_fore};
                                border: 1px solid #999;
                                box-shadow: 0 0 10px rgba(0,0,0,0.5);
                                font: {form_font};
                            ",
                            
                            // Title bar
                            div {
                                style: "
                                    height: 30px;
                                    background: linear-gradient(to bottom, #0078d4, #005a9e);
                                    color: {form_fore};
                                    padding: 4px 8px;
                                    display: flex;
                                    align-items: center;
                                    font-weight: bold;
                                ",
                                "{caption}"
                            }
                            
                            // Controls
                            for control in form.controls {
                                {
                                    let control_type = control.control_type;
                                    let x = control.bounds.x;
                                    let y = control.bounds.y;
                                    let w = control.bounds.width;
                                    let h = control.bounds.height;
                                    let name = control.name.clone();
                                    let name_clone = name.clone(); // For closure
                                    
                                    // Helper to get text/caption
                                    let caption = control.get_caption().map(|s| s.to_string()).unwrap_or(name.clone());
                                    let text = control.get_text().map(|s| s.to_string()).unwrap_or_default();
                                    let is_enabled = control.is_enabled();
                                    let is_visible = control.is_visible();
                                    
                                    let base_font = "font-family: 'Segoe UI', sans-serif; color: #0f172a;";
                                    let base_field_bg = "background: #f8fafc;";
                                    let base_button_bg = "background: linear-gradient(90deg, #2563eb, #1d4ed8); color: #f8fafc; border: 1px solid #1d4ed8;";
                                    let base_frame_border = "border: 1px solid #cbd5e1;";

                                    let back_color = control.get_back_color().map(|s| s.to_string());
                                    let fore_color = control.get_fore_color().map(|s| s.to_string());
                                    let font_family = control.get_font().map(|s| s.to_string());

                                    let mut style_font = base_font.to_string();
                                    if let Some(f) = &font_family {
                                        style_font = format!("font: {};", f);
                                    }
                                    let mut style_fore = String::new();
                                    if let Some(fc) = &fore_color { style_fore = format!("color: {};", fc); }
                                    let mut style_back = String::new();
                                    if let Some(bc) = &back_color { style_back = format!("background: {};", bc); }
                                    let button_bg = if let Some(bc) = &back_color {
                                        let color_part = if let Some(fc) = &fore_color { format!("color: {};", fc) } else { String::new() };
                                        format!("background: {}; {}; border: 1px solid #cbd5e1;", bc, color_part)
                                    } else {
                                        base_button_bg.to_string()
                                    };

                                    rsx! {
                                        if is_visible {
                                            div {
                                                style: "position: absolute; left: {x}px; top: {y}px; width: {w}px; height: {h}px;",
                                                
                                                {match control_type {
                                                    ControlType::Button => rsx! {
                                                        button {
                                                            style: "width: 100%; height: 100%; padding: 6px 10px; {button_bg} {style_font}; border-radius: 6px; box-shadow: 0 2px 4px rgba(37,99,235,0.12);",
                                                            disabled: !is_enabled,
                                                            onclick: move |_| handle_event(name_clone.clone(), "Click".to_string()),
                                                            "{caption}"
                                                        }
                                                    },
                                                    ControlType::TextBox => rsx! {
                                                        input {
                                                            style: "width: 100%; height: 100%; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            disabled: !is_enabled,
                                                            value: "{text}",
                                                            oninput: move |evt| {
                                                                 if let Some(frm) = runtime_form.write().as_mut() {
                                                                     if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                         ctrl.set_text(evt.value());
                                                                     }
                                                                 }
                                                                 handle_event(name_clone.clone(), "Change".to_string());
                                                            }
                                                        }
                                                    },
                                                    ControlType::Label => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; padding: 4px 2px; {style_font} {style_fore} {style_back};",
                                                            "{caption}"
                                                        }
                                                    },
                                                    ControlType::CheckBox => rsx! {
                                                        div {
                                                            style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                                            input { 
                                                                r#type: "checkbox",
                                                                disabled: !is_enabled,
                                                                checked: control.properties.get_int("Value").unwrap_or(0) == 1,
                                                                onclick: move |_| handle_event(name_clone.clone(), "Click".to_string())
                                                            }
                                                            span { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::RadioButton => rsx! {
                                                        div {
                                                            style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                                            input { 
                                                                r#type: "radio", 
                                                                name: "radio_group",
                                                                disabled: !is_enabled,
                                                                checked: control.properties.get_int("Value").unwrap_or(0) == 1,
                                                                onclick: move |_| handle_event(name_clone.clone(), "Click".to_string())
                                                            }
                                                            span { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::Frame => rsx! {
                                                        fieldset {
                                                            style: "width: 100%; height: 100%; {base_frame_border} margin: 0; padding: 0; border-radius: 8px; {style_back} {style_font} {style_fore};",
                                                            legend { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::ListBox => rsx! {
                                                        select {
                                                            style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            multiple: true,
                                                            disabled: !is_enabled,
                                                            onchange: move |_| handle_event(name_clone.clone(), "Click".to_string()),
                                                            {
                                                                // Prefer the List property; if empty, fall back to Text split by '|'
                                                                let mut items = control.get_list_items();
                                                                if items.is_empty() {
                                                                    let raw = text.clone();
                                                                    if !raw.is_empty() {
                                                                        items = raw.split('|').map(|s| s.trim().to_string()).collect();
                                                                    }
                                                                }

                                                                if items.is_empty() {
                                                                    rsx! { option { "(empty)" } }
                                                                } else {
                                                                    rsx! {
                                                                        for item in items {
                                                                            option { "{item}" }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::ComboBox => rsx! {
                                                        select {
                                                            style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            disabled: !is_enabled,
                                                            onchange: move |evt| {
                                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                        ctrl.set_text(evt.value());
                                                                    }
                                                                }
                                                                handle_event(name_clone.clone(), "Change".to_string());
                                                            },
                                                            {
                                                                let items = control.get_list_items();
                                                                if items.is_empty() {
                                                                    rsx! { option { value: "", "{text}" } }
                                                                } else {
                                                                    rsx! {
                                                                        for item in items {
                                                                            option { "{item}" }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::RichTextBox => rsx! {
                                                        {
                                                            // Use HTML property if set, otherwise fall back to Text
                                                            let html = control.properties.get_string("HTML")
                                                                .map(|s| s.to_string())
                                                                .or_else(|| control.get_text().map(|s| s.to_string()))
                                                                .unwrap_or_default();
                                                            let toolbar_visible = control.properties.get_bool("ToolbarVisible").unwrap_or(true);
                                                            let rtb_id = format!("rtb_{}", name_clone);
                                                            let rtb_id_bold = rtb_id.clone();
                                                            let rtb_id_italic = rtb_id.clone();
                                                            let rtb_id_underline = rtb_id.clone();
                                                            let rtb_id_ul = rtb_id.clone();
                                                            let rtb_id_ol = rtb_id.clone();
                                                            let rtb_id_mount = rtb_id.clone();
                                                            rsx! {
                                                                div {
                                                                    style: "width: 100%; height: 100%; display: flex; flex-direction: column; border: 1px inset #999; background: white; {style_back} {style_font} {style_fore};",
                                                                    // Toolbar (conditional)
                                                                    if toolbar_visible {
                                                                        div {
                                                                            style: "display: flex; gap: 2px; padding: 4px; background: #f0f0f0; border-bottom: 1px solid #ccc;",
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-weight: bold;",
                                                                                title: "Bold (Ctrl+B)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('bold', false, null); document.getElementById('{}').focus();", rtb_id_bold));
                                                                                },
                                                                                "B"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-style: italic;",
                                                                                title: "Italic (Ctrl+I)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('italic', false, null); document.getElementById('{}').focus();", rtb_id_italic));
                                                                                },
                                                                                "I"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; text-decoration: underline;",
                                                                                title: "Underline (Ctrl+U)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('underline', false, null); document.getElementById('{}').focus();", rtb_id_underline));
                                                                                },
                                                                                "U"
                                                                            }
                                                                            div { style: "width: 1px; background: #ccc; margin: 0 4px;" }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Bullet List",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('insertUnorderedList', false, null); document.getElementById('{}').focus();", rtb_id_ul));
                                                                                },
                                                                                "• List"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Numbered List",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('insertOrderedList', false, null); document.getElementById('{}').focus();", rtb_id_ol));
                                                                                },
                                                                                "1. List"
                                                                            }
                                                                        }
                                                                    }
                                                                    // ContentEditable div
                                                                    div {
                                                                        id: "{rtb_id}",
                                                                        contenteditable: if is_enabled { "true" } else { "false" },
                                                                        style: "flex: 1; padding: 8px; overflow: auto; outline: none; background: white; {style_back} {style_font} {style_fore};",
                                                                        dangerous_inner_html: "{html}",
                                                                        onmounted: move |_| {
                                                                            // Add keyboard shortcuts
                                                                            let _ = eval(&format!(r#"
                                                                                (function() {{
                                                                                    const editor = document.getElementById('{}');
                                                                                    if (editor) {{
                                                                                        editor.addEventListener('keydown', function(e) {{
                                                                                            if (e.ctrlKey || e.metaKey) {{
                                                                                                switch(e.key.toLowerCase()) {{
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
                                                                                                }}
                                                                                            }}
                                                                                        }});
                                                                                    }}
                                                                                }})();
                                                                            "#, rtb_id_mount));
                                                                        },
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::WebBrowser => rsx! {
                                                        {
                                                            let url = control.properties.get_string("URL").map(|s| s.to_string()).unwrap_or_else(|| "about:blank".to_string());
                                                            rsx! {
                                                                iframe {
                                                                    id: "{name_clone}",
                                                                    style: "width: 100%; height: 100%; border: 1px inset #999; background: white; {style_back};",
                                                                    src: "{url}",
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::ListView => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px inset #999; background: white; overflow: auto; {style_back} {style_font} {style_fore};",
                                                            // Placeholder content - in real impl, we'd iterate over Columns and Items
                                                            table {
                                                                style: "width: 100%; border-collapse: collapse; font-size: 12px; {style_font}; {style_fore};",
                                                                thead {
                                                                    tr {
                                                                        th { style: "border: 1px solid #ccc; padding: 2px; background: #f0f0f0; text-align: left; {style_fore}; {style_font};", "ColumnHeader" }
                                                                    }
                                                                }
                                                                tbody {
                                                                    tr {
                                                                        td { style: "border: 1px solid #eee; padding: 2px; {style_font}; {style_fore};", "ListItem" }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::TreeView => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px inset #999; background: white; overflow: auto; padding: 4px; {style_back} {style_font} {style_fore};",
                                                            div { style: "{style_font} {style_fore};", "Node0" }
                                                            div { style: "padding-left: 20px; {style_font} {style_fore};", "Node1" }
                                                        }
                                                    },
                                                    ControlType::DataGridView => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #999; background: #808080; padding: 1px; overflow: auto;",
                                                            table {
                                                                style: "width: 100%; background: white; border-collapse: separate; border-spacing: 0;",
                                                                thead {
                                                                    tr {
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px; width: 20px;", "" }
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px;", "A" }
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px;", "B" }
                                                                    }
                                                                }
                                                                tbody {
                                                                    tr {
                                                                        td { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center;", "1" }
                                                                        td { style: "border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; padding: 2px;", "" }
                                                                        td { style: "border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; padding: 2px;", "" }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::Panel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #ccc; overflow: hidden; {style_back} {style_font} {style_fore};",
                                                            // We rely on child controls having correct coordinates relative to form (flattened) 
                                                            // OR if we support nesting, we'd need to render children here.
                                                            // Current architecture flattens controls onto the form, so this is just a background.
                                                        }
                                                    },
                                                    _ => rsx! {
                                                        div { style: "border: 1px dotted red;", "Unknown Control" }
                                                    }
                                                }}
                                            }
                                        }
                                    }
                                }
                            }
                            // ... MsgBox ...
                            {if msgbox_visible {
                                rsx! {
                                    div {
                                        style: "
                                            position: absolute;
                                            top: 0; left: 0; right: 0; bottom: 0;
                                            background: rgba(0,0,0,0.5);
                                            display: flex;
                                            align-items: center;
                                            justify-content: center;
                                            z-index: 1000;
                                        ",
                                        div {
                                            style: "
                                                background: #f0f0f0;
                                                border: 1px solid #999;
                                                box-shadow: 0 4px 16px rgba(0,0,0,0.5);
                                                min-width: 200px;
                                                display: flex;
                                                flex-direction: column;
                                            ",
                                            // Title
                                            div {
                                                style: "background: #0078d4; color: white; padding: 4px 8px; font-weight: bold;",
                                                "Project1" // Default title 
                                            }
                                            // Content
                                            div {
                                                style: "padding: 20px; text-align: center; color: black;",
                                                "{msgbox_text}"
                                            }
                                            // Button
                                            div {
                                                style: "padding: 10px; display: flex; justify-content: center;",
                                                button {
                                                    style: "padding: 4px 20px;",
                                                    onclick: move |_| msgbox_content.set(None),
                                                    "OK"
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                rsx! {}
                            }}
                        }
                    }
                } else {
                    rsx! { div { "Loading..." } }
                }
            }
        }
    }
}
