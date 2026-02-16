use dioxus::prelude::*;
use dioxus::desktop::use_asset_handler;
use wry::http::Response;
use vybe_forms::{ControlType, Form, EventType};
use vybe_project::Project;
use vybe_runtime::{Interpreter, RuntimeSideEffect, Value, ObjectData, ConsoleMessage};
use std::collections::HashMap;
use std::cell::RefCell;
use std::rc::Rc;
use std::sync::mpsc;
use vybe_parser::parse_program;
use crate::runner::LAUNCH_PROJECT;

// ---------------------------------------------------------------------------
// Console color helpers
// ---------------------------------------------------------------------------

/// Map .NET ConsoleColor value (0-15) to a CSS color string.
fn console_color_to_css(color: i32) -> &'static str {
    match color {
        0  => "#0c0c0c",  // Black
        1  => "#0037da",  // DarkBlue
        2  => "#13a10e",  // DarkGreen
        3  => "#3a96dd",  // DarkCyan
        4  => "#c50f1f",  // DarkRed
        5  => "#881798",  // DarkMagenta
        6  => "#c19c00",  // DarkYellow
        7  => "#cccccc",  // Gray
        8  => "#767676",  // DarkGray
        9  => "#3b78ff",  // Blue
        10 => "#16c60c",  // Green
        11 => "#61d6d6",  // Cyan
        12 => "#e74856",  // Red
        13 => "#b4009e",  // Magenta
        14 => "#f9f1a5",  // Yellow
        15 => "#f2f2f2",  // White
        _  => "#cccccc",  // default Gray
    }
}

/// Escape HTML special characters.
fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
}

// ---------------------------------------------------------------------------
// Public context wrapper – both the editor and the CLI provide this via
// `use_context_provider`.  It is the single source-of-truth the FormRunner
// reads from, so there is exactly one code path.
// ---------------------------------------------------------------------------

/// Wrapper around a `Signal<Option<Project>>` that the FormRunner reads.
/// Both the IDE and the standalone shell provide this via `use_context_provider`.
#[derive(Clone, Copy)]
pub struct RuntimeProject {
    pub project: Signal<Option<Project>>,
    /// Set to `true` by the FormRunner when a console (Sub Main) project
    /// finishes executing.  The editor watches this to leave run-mode
    /// automatically; the standalone shell can ignore it.
    pub finished: Signal<bool>,
}

/// Top-level Dioxus component for the standalone shell.
#[component]
fn ShellApp() -> Element {
    let project = LAUNCH_PROJECT
        .with(|cell| cell.borrow().clone())
        .expect("LAUNCH_PROJECT must be set before launching");

    use_context_provider(|| RuntimeProject {
        project: Signal::new(Some(project)),
        finished: Signal::new(false),
    });

    rsx! { FormRunner {} }
}

#[derive(PartialEq, Props, Clone)]
struct ControlTreeProps {
    form: Form,
    parent_id: Option<uuid::Uuid>,
    interpreter: Signal<Option<Interpreter>>,
    wb_html: Signal<HashMap<String, String>>,
    runtime_form: Signal<Option<Form>>,
    #[props(into)]
    on_handle_event: EventHandler<(String, String, Option<vybe_runtime::EventData>)>,
}

#[component]
fn ControlTree(props: ControlTreeProps) -> Element {
    let form = &props.form;
    let parent_id = props.parent_id;
    let interpreter = props.interpreter;
    let wb_html = props.wb_html;
    let mut runtime_form = props.runtime_form;
    let on_handle_event = props.on_handle_event;

    // Filter controls for this level
    // We must clone because we can't return references to form.controls in the iterator easily
    let children: Vec<_> = form.controls.iter()
        .filter(|c| c.parent_id == parent_id && !c.control_type.is_non_visual())
        .cloned()
        .collect();
    
    // Sort logic if needed (currently reliance on vector order is default)
    
    if children.is_empty() {
        return rsx! {};
    }

    rsx! {
        for control in children {
            {
                let control_type = control.control_type.clone();
                let x = control.bounds.x;
                let y = control.bounds.y;
                let w = control.bounds.width;
                let h = control.bounds.height;
                let name = control.name.clone();
                let id = control.id;
                
                // Local handle_event wrapper to match existing code style
                let handle_event = {
                    let handler = on_handle_event;
                    move |n: String, e: String, d: Option<vybe_runtime::EventData>| {
                        handler.call((n, e, d))
                    }
                };

                let text = control.get_text().map(|s| s.to_string()).unwrap_or_else(|| name.clone());
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
                if let Some(fc) = &fore_color {
                    style_fore = format!("color: {};", fc);
                }
                let mut style_back = String::new();
                if let Some(bc) = &back_color {
                    style_back = format!("background: {};", bc);
                }
                let button_bg = if let Some(bc) = &back_color {
                    let color_part = if let Some(fc) = &fore_color {
                        format!("color: {};", fc)
                    } else {
                        String::new()
                    };
                    format!("background: {}; {}; border: 1px solid #cbd5e1;", bc, color_part)
                } else {
                    base_button_bg.to_string()
                };

                let parent = parent_id.and_then(|pid| form.controls.iter().find(|c| c.id == pid));
                let is_flow_or_table = parent.map(|p| matches!(p.control_type, 
                    ControlType::FlowLayoutPanel | ControlType::TableLayoutPanel |
                    ControlType::ToolStrip | ControlType::MenuStrip | ControlType::StatusStrip
                )).unwrap_or(false);

                // Compute Dock-based positioning
                let dock_val = control.properties.get_int("Dock").unwrap_or(0);
                
                let wrapper_style = if is_flow_or_table {
                    // Static positioning for flow/table layouts
                    format!("position: relative; width: {}px; height: {}px; margin: 2px;", w, h)
                } else {
                    match dock_val {
                        1 => "position: absolute; z-index: 10; left: 0; top: 0; width: 100%; outline: none;".to_string(), // DockStyle.Top
                        2 => "position: absolute; z-index: 10; left: 0; bottom: 0; width: 100%; outline: none;".to_string(), // DockStyle.Bottom
                        3 => "position: absolute; z-index: 10; left: 0; top: 0; height: 100%; outline: none;".to_string(), // DockStyle.Left
                        4 => "position: absolute; z-index: 10; right: 0; top: 0; height: 100%; outline: none;".to_string(), // DockStyle.Right
                        5 => "position: absolute; z-index: 10; left: 0; top: 0; width: 100%; height: 100%; outline: none;".to_string(), // DockStyle.Fill
                        _ => format!("position: absolute; z-index: 10; left: {}px; top: {}px; width: {}px; height: {}px; outline: none;", x, y, w, h),
                    }
                };

                let name_clone = name.clone();
                let name_focusin = name.clone();
                let name_focusout = name.clone();
                
                println!("[DEBUG] Rendering control '{}' (ID={}) inside parent {:?}. Style: [{}] Visible: {}", 
                    name_clone, control.id, parent_id, wrapper_style, is_visible);

                rsx! {
                    if is_visible {
                        div {
                            style: "{wrapper_style}",
                            // Event bubbling / capture
                            onfocusin: move |_evt: FocusEvent| {
                                handle_event(name_focusin.clone(), "GotFocus".to_string(), None);
                            },
                            onfocusout: move |_evt: FocusEvent| {
                                handle_event(name_focusout.clone(), "LostFocus".to_string(), None);
                            },

                            {match control_type {
                                ControlType::Panel => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; position: relative; border: 1px solid #ccc; {style_back} {style_font} {style_fore};",
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::Frame => rsx! {
                                    fieldset {
                                        style: "width: 100%; height: 100%; position: relative; {base_frame_border} margin: 0; padding: 0; border-radius: 8px; {style_back} {style_font} {style_fore};",
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        legend { "{text}" }
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::Button => rsx! {
                                    button {
                                        style: "width: 100%; height: 100%; padding: 6px 10px; {button_bg} {style_font}; border-radius: 6px; box-shadow: 0 2px 4px rgba(37,99,235,0.12);",
                                        disabled: !is_enabled,
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        "{text}"
                                    }
                                },
                                ControlType::Label => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; padding: 4px 2px; {style_font} {style_fore} {style_back};",
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        "{text}"
                                    }
                                },
                                ControlType::TextBox => {
                                    let is_multiline = control.properties.get_bool("Multiline").unwrap_or(false);
                                    let is_readonly = control.properties.get_bool("ReadOnly").unwrap_or(false);
                                    if is_multiline {
                                        rsx! {
                                            textarea {
                                                style: "width: 100%; height: 100%; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; resize: none; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                disabled: !is_enabled,
                                                readonly: is_readonly,
                                                value: "{text}",
                                                oninput: move |evt| {
                                                    if let Some(frm) = runtime_form.write().as_mut() {
                                                        if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                            ctrl.set_text(evt.value());
                                                        }
                                                    }
                                                    handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                                    handle_event(name_clone.clone(), "Change".to_string(), None);
                                                }
                                            }
                                        }
                                    } else {
                                        rsx! {
                                            input {
                                                style: "width: 100%; height: 100%; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                disabled: !is_enabled,
                                                readonly: is_readonly,
                                                value: "{text}",
                                                oninput: move |evt| {
                                                    if let Some(frm) = runtime_form.write().as_mut() {
                                                        if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                            ctrl.set_text(evt.value());
                                                        }
                                                    }
                                                    handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                                    handle_event(name_clone.clone(), "Change".to_string(), None);
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::CheckBox => rsx! {
                                    div {
                                        style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                        input {
                                            r#type: "checkbox",
                                            disabled: !is_enabled,
                                            checked: control.properties.get_bool("Checked").unwrap_or(false) || control.properties.get_int("Value").unwrap_or(0) == 1,
                                            onclick: move |evt: MouseEvent| {
                                                // Toggle state
                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                        let was_checked = ctrl.properties.get_bool("Checked").unwrap_or(false) || ctrl.properties.get_int("Value").unwrap_or(0) == 1;
                                                        let now_checked = !was_checked;
                                                        ctrl.properties.set("Checked", now_checked);
                                                        use vybe_forms::properties::PropertyValue;
                                                        let int_val = if now_checked { 1 } else { 0 };
                                                        ctrl.properties.set_raw("Value", PropertyValue::Integer(int_val));
                                                        ctrl.properties.set_raw("CheckState", PropertyValue::Integer(int_val));
                                                    }
                                                }
                                                handle_event(name_clone.clone(), "CheckedChanged".to_string(), None);
                                                let data = vybe_runtime::EventData::Mouse {
                                                    button: 0x100000, clicks: 1,
                                                    x: evt.client_coordinates().x as i32,
                                                    y: evt.client_coordinates().y as i32,
                                                    delta: 0,
                                                };
                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                            }
                                        }
                                        span { "{text}" }
                                    }
                                },
                                ControlType::RadioButton => rsx! {
                                    div {
                                        style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                        input {
                                            r#type: "radio",
                                            name: "radio_group",
                                            disabled: !is_enabled,
                                            checked: control.properties.get_bool("Checked").unwrap_or(false) || control.properties.get_int("Value").unwrap_or(0) == 1,
                                            onclick: move |evt: MouseEvent| {
                                                // Toggle state
                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                        ctrl.properties.set("Checked", true);
                                                        use vybe_forms::properties::PropertyValue;
                                                        ctrl.properties.set_raw("Value", PropertyValue::Integer(1));
                                                    }
                                                }
                                                handle_event(name_clone.clone(), "CheckedChanged".to_string(), None);
                                                let data = vybe_runtime::EventData::Mouse {
                                                    button: 0x100000, clicks: 1,
                                                    x: evt.client_coordinates().x as i32,
                                                    y: evt.client_coordinates().y as i32,
                                                    delta: 0,
                                                };
                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                            }
                                        }
                                        span { "{text}" }
                                    }
                                },
                                ControlType::ListBox => rsx! {
                                    select {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                        multiple: true,
                                        disabled: !is_enabled,
                                        onchange: move |evt| {
                                            if let Some(frm) = runtime_form.write().as_mut() {
                                                if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                    ctrl.set_text(evt.value());
                                                    let items = ctrl.get_list_items();
                                                    if let Some(idx) = items.iter().position(|i| *i == evt.value()) {
                                                        use vybe_forms::properties::PropertyValue;
                                                        ctrl.properties.set_raw("SelectedIndex", PropertyValue::Integer(idx as i32));
                                                    }
                                                }
                                            }
                                            handle_event(name_clone.clone(), "SelectedIndexChanged".to_string(), None);
                                            handle_event(name_clone.clone(), "Click".to_string(), None);
                                        },
                                        {
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
                                                    let items = ctrl.get_list_items();
                                                    if let Some(idx) = items.iter().position(|i| *i == evt.value()) {
                                                        use vybe_forms::properties::PropertyValue;
                                                        ctrl.properties.set_raw("SelectedIndex", PropertyValue::Integer(idx as i32));
                                                    }
                                                }
                                            }
                                            handle_event(name_clone.clone(), "SelectedIndexChanged".to_string(), None);
                                            handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                            handle_event(name_clone.clone(), "Change".to_string(), None);
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
                                        let html = control.properties.get_string("HTML")
                                            .map(|s| s.to_string())
                                            .or_else(|| control.get_text().map(|s| s.to_string()))
                                            .unwrap_or_default();
                                        let rtb_id = format!("rtb_{}", name_clone);
                                        rsx! {
                                            div {
                                                style: "width: 100%; height: 100%; display: flex; flex-direction: column; border: 1px inset #999; background: white; {style_back} {style_font} {style_fore};",
                                                div {
                                                    id: "{rtb_id}",
                                                    contenteditable: if is_enabled { "true" } else { "false" },
                                                    style: "flex: 1; padding: 8px; overflow: auto; outline: none; background: white; {style_back} {style_font} {style_fore};",
                                                    dangerous_inner_html: "{html}",
                                                    oninput: move |_| {
                                                        handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                                        handle_event(name_clone.clone(), "Change".to_string(), None);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::WebBrowser => {
                                    let content = wb_html.read()
                                        .get(&name_clone.to_lowercase())
                                        .cloned()
                                        .or_else(|| control.properties.get_string("HTML").map(|s| s.to_string()))
                                        .unwrap_or_default();
                                    if content.is_empty() {
                                        rsx! {
                                            div {
                                                id: "wb_{name_clone}",
                                                style: "width: 100%; height: 100%; border: 1px inset #999; background: white; {style_back};",
                                            }
                                        }
                                    } else {
                                        rsx! {
                                            div {
                                                id: "wb_{name_clone}",
                                                style: "width: 100%; height: 100%; border: 1px inset #999; background: white; overflow: hidden; {style_back};",
                                                dangerous_inner_html: "{content}",
                                            }
                                        }
                                    }
                                },
                                ControlType::ListView => {
                                    let lv_items = control.get_list_items();
                                    let lv_selected = control.properties.get_int("SelectedIndex").unwrap_or(-1);
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; border: 1px inset #999; overflow: auto; {base_field_bg} {style_back} {style_font} {style_fore};",
                                            table {
                                                style: "width: 100%; border-collapse: collapse; font-size: 12px;",
                                                thead { tr { th { style: "border-bottom: 2px solid #aaa; padding: 3px 6px; background: #f0f0f0; text-align: left;", "Item" } } }
                                                tbody {
                                                    if lv_items.is_empty() {
                                                        tr { td { style: "padding: 4px 6px; color: #aaa; font-style: italic;", "(empty)" } }
                                                    } else {
                                                        for (idx, item) in lv_items.iter().enumerate() {
                                                            {
                                                                let item_str = item.clone();
                                                                let item_name = name_clone.clone();
                                                                let item_idx = idx as i32;
                                                                let row_style = if item_idx == lv_selected {
                                                                    "background: #cce5ff; cursor: pointer; border-bottom: 1px solid #eee; padding: 3px 6px;"
                                                                } else {
                                                                    "cursor: pointer; border-bottom: 1px solid #eee; padding: 3px 6px;"
                                                                };
                                                                rsx! {
                                                                    tr {
                                                                        key: "{idx}",
                                                                        onclick: move |_| {
                                                                            if let Some(frm) = runtime_form.write().as_mut() {
                                                                                if let Some(ctrl) = frm.get_control_by_name_mut(&item_name) {
                                                                                    ctrl.properties.set("SelectedIndex", item_idx);
                                                                                }
                                                                            }
                                                                            handle_event(item_name.clone(), "SelectedIndexChanged".to_string(), None);
                                                                            handle_event(item_name.clone(), "ItemActivate".to_string(), None);
                                                                        },
                                                                        td { style: "{row_style}", "{item_str}" }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::UserControl => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; position: relative; border: 1px dashed #94a3b8; {style_back};",
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::FlowLayoutPanel => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #e2e8f0; display: flex; flex-wrap: wrap; align-content: flex-start; padding: 2px; {style_back}",
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::TableLayoutPanel => {
                                    let tlp_cols = control.properties.get_int("ColumnCount").unwrap_or(2);
                                    let tlp_rows = control.properties.get_int("RowCount").unwrap_or(2);
                                    let grid_cols = format!("repeat({}, 1fr)", tlp_cols);
                                    let grid_rows = format!("repeat({}, 1fr)", tlp_rows);
                                    // We render children here so they fill the grid
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; border: 1px solid #e2e8f0; display: grid; grid-template-columns: {grid_cols}; grid-template-rows: {grid_rows}; {style_back}",
                                            ControlTree {
                                                form: form.clone(),
                                                parent_id: Some(id),
                                                interpreter: interpreter,
                                                wb_html: wb_html,
                                                runtime_form: runtime_form,
                                                on_handle_event: on_handle_event
                                            }
                                        }
                                    }
                                },
                                ControlType::TabControl => {
                                    let tab_items = control.get_list_items();
                                    let selected_tab = control.properties.get_int("SelectedIndex").unwrap_or(0);
                                    let tabs: Vec<String> = if tab_items.is_empty() { vec!["Tab 1".to_string()] } else { tab_items };
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; border: 1px solid #adb5bd; display: flex; flex-direction: column;",
                                            div {
                                                style: "display: flex; background: #e9ecef; border-bottom: 1px solid #adb5bd;",
                                                for (ti, tab_label) in tabs.iter().enumerate() {
                                                    {
                                                        let tl = tab_label.clone();
                                                        let tab_name = name_clone.clone();
                                                        let is_active = ti as i32 == selected_tab;
                                                        let tab_style = if is_active {
                                                            "padding: 4px 12px; background: white; border: 1px solid #adb5bd; border-bottom: none; cursor: pointer; font-size: 12px; font-weight: bold;"
                                                        } else {
                                                            "padding: 4px 12px; background: #e9ecef; border: 1px solid transparent; cursor: pointer; font-size: 12px;"
                                                        };
                                                        rsx! {
                                                            div {
                                                                style: "{tab_style}",
                                                                onclick: move |_| {
                                                                    if let Some(frm) = runtime_form.write().as_mut() {
                                                                        if let Some(ctrl) = frm.get_control_by_name_mut(&tab_name) {
                                                                            use vybe_forms::properties::PropertyValue;
                                                                            ctrl.properties.set_raw("SelectedIndex", PropertyValue::Integer(ti as i32));
                                                                        }
                                                                    }
                                                                    handle_event(tab_name.clone(), "SelectedIndexChanged".to_string(), None);
                                                                },
                                                                "{tl}"
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            div {
                                                style: "flex: 1; padding: 8px; background: white;",
                                            }
                                        }
                                    }
                                },
                                ControlType::BindingNavigator => rsx! {
                                    {
                                        let nav_first = name.clone();
                                        let nav_prev = name.clone();
                                        let nav_next = name.clone();
                                        let nav_last = name.clone();
                                        let nav_add = name.clone();
                                        let nav_del = name.clone();
                                        rsx! {
                                            div {
                                                style: "width: 100%; height: 100%; display: flex; align-items: center; gap: 2px; background: #f0f0f0; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Move first",
                                                    onclick: move |_| handle_event(nav_first.clone(), "MoveFirst".to_string(), None),
                                                    "⏮"
                                                }
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Move previous",
                                                    onclick: move |_| handle_event(nav_prev.clone(), "MovePrevious".to_string(), None),
                                                    "◀"
                                                }
                                                span {
                                                    style: "padding: 0 4px; min-width: 40px; text-align: center; border: 1px solid #ccc; background: white;",
                                                    "{text}"
                                                }
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Move next",
                                                    onclick: move |_| handle_event(nav_next.clone(), "MoveNext".to_string(), None),
                                                    "▶"
                                                }
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Move last",
                                                    onclick: move |_| handle_event(nav_last.clone(), "MoveLast".to_string(), None),
                                                    "⏭"
                                                }
                                                div { style: "width: 1px; height: 16px; background: #aaa; margin: 0 2px;" }
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Add new",
                                                    onclick: move |_| handle_event(nav_add.clone(), "AddNew".to_string(), None),
                                                    "➕"
                                                }
                                                button {
                                                    style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                    title: "Delete",
                                                    onclick: move |_| handle_event(nav_del.clone(), "Delete".to_string(), None),
                                                    "❌"
                                                }
                                                ControlTree {
                                                    form: form.clone(),
                                                    parent_id: Some(id),
                                                    interpreter: interpreter,
                                                    wb_html: wb_html,
                                                    runtime_form: runtime_form,
                                                    on_handle_event: on_handle_event
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::PictureBox => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #e2e8f0; overflow: hidden; {style_back};",
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::ProgressBar => {
                                    let pb_val = control.properties.get_int("Value").unwrap_or(0);
                                    let pb_max = control.properties.get_int("Maximum").unwrap_or(100);
                                    let pb_pct = if pb_max > 0 { (pb_val as f64 / pb_max as f64 * 100.0) as i32 } else { 0 };
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; background: #e9ecef; border: 1px solid #adb5bd; overflow: hidden; border-radius: 4px;",
                                            div {
                                                style: "height: 100%; background: linear-gradient(180deg, #5cb85c, #4cae4c); width: {pb_pct}%; transition: width 0.3s;",
                                            }
                                        }
                                    }
                                },
                                ControlType::NumericUpDown => {
                                    let nud_val = control.properties.get_int("Value").unwrap_or(0);
                                    let nud_min = control.properties.get_int("Minimum").unwrap_or(0);
                                    let nud_max = control.properties.get_int("Maximum").unwrap_or(100);
                                    let nud_val_str = nud_val.to_string();
                                    let nud_min_str = nud_min.to_string();
                                    let nud_max_str = nud_max.to_string();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; border: 1px solid #adb5bd; border-radius: 6px; overflow: hidden; {style_back}",
                                            input {
                                                r#type: "number",
                                                style: "flex: 1; border: none; padding: 2px 6px; {style_font} {style_fore} outline: none; background: transparent;",
                                                disabled: !is_enabled,
                                                min: "{nud_min_str}",
                                                max: "{nud_max_str}",
                                                value: "{nud_val_str}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Ok(v) = evt.value().parse::<i32>() {
                                                        if let Some(frm) = runtime_form.write().as_mut() {
                                                            if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                ctrl.properties.set("Value", v);
                                                                ctrl.set_text(v.to_string());
                                                            }
                                                        }
                                                        handle_event(name_clone.clone(), "ValueChanged".to_string(), None);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::MenuStrip | ControlType::ContextMenuStrip => {
                                    let menu_items = control.get_list_items();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; background: #f0f0f0; border-bottom: 1px solid #ccc; display: flex; align-items: center; padding: 0 4px; {style_font} font-size: 12px;",
                                            if menu_items.is_empty() {
                                                span {
                                                    style: "padding: 2px 8px; cursor: pointer;",
                                                    onclick: move |_| handle_event(name_clone.clone(), "Click".to_string(), None),
                                                    "File"
                                                }
                                            } else {
                                                for item in menu_items {
                                                    {
                                                        let item_text = item.clone();
                                                        let menu_name = name_clone.clone();
                                                        rsx! {
                                                            span {
                                                                style: "padding: 2px 8px; cursor: pointer;",
                                                                onclick: move |_| {
                                                                    handle_event(menu_name.clone(), "ItemClicked".to_string(), None);
                                                                },
                                                                "{item_text}"
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            ControlTree {
                                                form: form.clone(),
                                                parent_id: Some(id),
                                                interpreter: interpreter,
                                                wb_html: wb_html,
                                                runtime_form: runtime_form,
                                                on_handle_event: on_handle_event
                                            }
                                        }
                                    }
                                },
                                ControlType::StatusStrip => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; background: #007acc; border-top: 1px solid #005a9e; display: flex; align-items: center; padding: 0 8px; font-size: 12px; color: white; {style_font}",
                                        "{text}"
                                        ControlTree {
                                            form: form.clone(),
                                            parent_id: Some(id),
                                            interpreter: interpreter,
                                            wb_html: wb_html,
                                            runtime_form: runtime_form,
                                            on_handle_event: on_handle_event
                                        }
                                    }
                                },
                                ControlType::DateTimePicker => {
                                    let dtp_format = control.properties.get_string("Format").map(|s| s.to_string()).unwrap_or_else(|| "Long".to_string());
                                    let input_type = if dtp_format == "Time" { "time" } else { "date" };
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; align-items: center; {style_back} {style_font} {style_fore}",
                                            input {
                                                r#type: "{input_type}",
                                                style: "width: 100%; height: 100%; border: 1px solid #adb5bd; padding: 2px 6px; border-radius: 6px; {style_font} {style_fore} {style_back} outline: none; cursor: pointer;",
                                                disabled: !is_enabled,
                                                value: "{text}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Some(frm) = runtime_form.write().as_mut() {
                                                        if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                            ctrl.set_text(evt.value());
                                                            ctrl.properties.set("Value", evt.value());
                                                        }
                                                    }
                                                    handle_event(name_clone.clone(), "ValueChanged".to_string(), None);
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::LinkLabel => rsx! {
                                    a {
                                        style: "width: 100%; height: 100%; display: flex; align-items: center; {style_font} font-size: 12px; color: #0066cc; text-decoration: underline; cursor: pointer;",
                                        onclick: move |evt: MouseEvent| {
                                            let data = vybe_runtime::EventData::Mouse {
                                                button: 0x100000, clicks: 1,
                                                x: evt.client_coordinates().x as i32,
                                                y: evt.client_coordinates().y as i32,
                                                delta: 0,
                                            };
                                            handle_event(name_clone.clone(), "LinkClicked".to_string(), Some(data.clone()));
                                            handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                        },
                                        "{text}"
                                    }
                                },
                                ControlType::ToolStrip => {
                                    let ts_items = control.get_list_items();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; background: #f0f0f0; border-bottom: 1px solid #ccc; display: flex; align-items: center; gap: 1px; padding: 2px 4px; {style_font} font-size: 11px;",
                                            if ts_items.is_empty() {
                                                span {
                                                    style: "padding: 2px 6px; background: #e8e8e8; border: 1px solid #ccc; cursor: pointer;",
                                                    onclick: move |_| handle_event(name_clone.clone(), "ButtonClick".to_string(), None),
                                                    "Button1"
                                                }
                                            } else {
                                                for it in ts_items {
                                                    {
                                                        let it_text = it.clone();
                                                        let ts_name = name_clone.clone();
                                                        rsx! {
                                                            span {
                                                                style: "padding: 2px 6px; background: #e8e8e8; border: 1px solid #ccc; cursor: pointer;",
                                                                onclick: move |_| {
                                                                    handle_event(ts_name.clone(), "ItemClicked".to_string(), None);
                                                                },
                                                                "{it_text}"
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            ControlTree {
                                                form: form.clone(),
                                                parent_id: Some(id),
                                                interpreter: interpreter,
                                                wb_html: wb_html,
                                                runtime_form: runtime_form,
                                                on_handle_event: on_handle_event
                                            }
                                        }
                                    }
                                },
                                ControlType::TrackBar => {
                                    let tb_val = control.properties.get_int("Value").unwrap_or(0);
                                    let tb_min = control.properties.get_int("Minimum").unwrap_or(0);
                                    let tb_max = control.properties.get_int("Maximum").unwrap_or(10);
                                    let tb_val_str = tb_val.to_string();
                                    let tb_min_str = tb_min.to_string();
                                    let tb_max_str = tb_max.to_string();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; align-items: center; padding: 4px; {style_back}",
                                            input {
                                                r#type: "range",
                                                style: "width: 100%; cursor: pointer;",
                                                disabled: !is_enabled,
                                                min: "{tb_min_str}",
                                                max: "{tb_max_str}",
                                                value: "{tb_val_str}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Ok(v) = evt.value().parse::<i32>() {
                                                        if let Some(frm) = runtime_form.write().as_mut() {
                                                            if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                ctrl.properties.set("Value", v);
                                                            }
                                                        }
                                                        handle_event(name_clone.clone(), "Scroll".to_string(), None);
                                                        handle_event(name_clone.clone(), "ValueChanged".to_string(), None);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::MaskedTextBox => {
                                    let mask = control.properties.get_string("Mask").map(|s| s.to_string()).unwrap_or_default();
                                    let prompt_char = control.properties.get_string("PromptChar").map(|s| s.to_string()).unwrap_or_else(|| "_".to_string());
                                    let placeholder = if mask.is_empty() { String::new() } else { mask.chars().map(|c| if c == '0' || c == '9' || c == '#' { prompt_char.chars().next().unwrap_or('_') } else { c }).collect::<String>() };
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                            input {
                                                r#type: "text",
                                                style: "width: 100%; height: 100%; border: 1px solid #adb5bd; padding: 2px 6px; border-radius: 6px; {style_font} {style_fore} {style_back} outline: none;",
                                                disabled: !is_enabled,
                                                value: "{text}",
                                                placeholder: "{placeholder}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Some(frm) = runtime_form.write().as_mut() {
                                                        if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                            ctrl.set_text(evt.value());
                                                        }
                                                    }
                                                    handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::SplitContainer => {
                                    let sc_orient = control.properties.get_string("Orientation").map(|s| s.to_string()).unwrap_or_else(|| "Vertical".to_string());
                                    let flex_dir = if sc_orient == "Horizontal" { "column" } else { "row" };
                                    let splitter_style = if sc_orient == "Horizontal" { "height: 4px; background: #d0d0d0; cursor: ns-resize; flex-shrink: 0;" } else { "width: 4px; background: #d0d0d0; cursor: ew-resize; flex-shrink: 0;" };
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; flex-direction: {flex_dir}; border: 1px solid #adb5bd; {style_back}",
                                            div { style: "flex: 1; overflow: hidden;" }
                                            div { style: "{splitter_style}" }
                                            div { style: "flex: 1; overflow: hidden;" }
                                        }
                                    }
                                },
                                ControlType::MonthCalendar => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; display: flex; align-items: stretch;",
                                        input {
                                            r#type: "date",
                                            style: "width: 100%; height: 100%; border: 1px solid #adb5bd; border-radius: 6px; padding: 4px 8px; {style_font} {style_fore} {style_back} outline: none; cursor: pointer;",
                                            disabled: !is_enabled,
                                            value: "{text}",
                                            oninput: move |evt: FormEvent| {
                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                        ctrl.set_text(evt.value());
                                                        ctrl.properties.set("SelectionStart", evt.value());
                                                    }
                                                }
                                                handle_event(name_clone.clone(), "DateChanged".to_string(), None);
                                            }
                                        }
                                    }
                                },
                                ControlType::HScrollBar => {
                                    let hs_val = control.properties.get_int("Value").unwrap_or(0);
                                    let hs_min = control.properties.get_int("Minimum").unwrap_or(0);
                                    let hs_max = control.properties.get_int("Maximum").unwrap_or(100);
                                    let hs_val_str = hs_val.to_string();
                                    let hs_min_str = hs_min.to_string();
                                    let hs_max_str = hs_max.to_string();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                            input {
                                                r#type: "range",
                                                style: "width: 100%; height: 100%; cursor: pointer;",
                                                disabled: !is_enabled,
                                                min: "{hs_min_str}",
                                                max: "{hs_max_str}",
                                                value: "{hs_val_str}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Ok(v) = evt.value().parse::<i32>() {
                                                        if let Some(frm) = runtime_form.write().as_mut() {
                                                            if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                ctrl.properties.set("Value", v);
                                                            }
                                                        }
                                                        handle_event(name_clone.clone(), "Scroll".to_string(), None);
                                                        handle_event(name_clone.clone(), "ValueChanged".to_string(), None);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::VScrollBar => {
                                    let vs_val = control.properties.get_int("Value").unwrap_or(0);
                                    let vs_min = control.properties.get_int("Minimum").unwrap_or(0);
                                    let vs_max = control.properties.get_int("Maximum").unwrap_or(100);
                                    let vs_val_str = vs_val.to_string();
                                    let vs_min_str = vs_min.to_string();
                                    let vs_max_str = vs_max.to_string();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                            input {
                                                r#type: "range",
                                                style: "width: 17px; height: 100%; writing-mode: vertical-lr; direction: rtl; cursor: pointer;",
                                                disabled: !is_enabled,
                                                min: "{vs_min_str}",
                                                max: "{vs_max_str}",
                                                value: "{vs_val_str}",
                                                oninput: move |evt: FormEvent| {
                                                    if let Ok(v) = evt.value().parse::<i32>() {
                                                        if let Some(frm) = runtime_form.write().as_mut() {
                                                            if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                ctrl.properties.set("Value", v);
                                                            }
                                                        }
                                                        handle_event(name_clone.clone(), "Scroll".to_string(), None);
                                                        handle_event(name_clone.clone(), "ValueChanged".to_string(), None);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::CheckedListBox => {
                                    let items = control.get_list_items();
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; overflow-y: auto; border: 1px solid #cbd5e1; border-radius: 4px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                            for (idx, item) in items.iter().enumerate() {
                                                {
                                                    let item_name = name_clone.clone();
                                                    let item_idx = idx;
                                                    let item_str = item.clone();
                                                    rsx! {
                                                        div {
                                                            key: "{item_idx}",
                                                            style: "display: flex; align-items: center; gap: 6px; padding: 2px 6px; cursor: pointer;",
                                                            onclick: move |_| {
                                                                handle_event(item_name.clone(), "ItemCheck".to_string(), None);
                                                                handle_event(item_name.clone(), "SelectedIndexChanged".to_string(), None);
                                                            },
                                                            input {
                                                                r#type: "checkbox",
                                                                disabled: !is_enabled,
                                                            }
                                                            span { "{item_str}" }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::DomainUpDown => rsx! {
                                    div {
                                        style: "display: flex; align-items: center; width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 4px; {base_field_bg} {style_back};",
                                        input {
                                            style: "flex: 1; border: none; outline: none; background: transparent; padding: 0 4px; {style_font} {style_fore};",
                                            disabled: !is_enabled,
                                            value: "{text}",
                                            oninput: move |evt| {
                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                        ctrl.set_text(evt.value());
                                                    }
                                                }
                                                handle_event(name_clone.clone(), "SelectedItemChanged".to_string(), None);
                                            }
                                        }
                                        div {
                                            style: "display: flex; flex-direction: column; border-left: 1px solid #cbd5e1;",
                                            button { style: "padding: 0 4px; border: none; background: #f1f5f9; cursor: pointer; font-size: 8px; line-height: 1;", disabled: !is_enabled, "▲" }
                                            button { style: "padding: 0 4px; border: none; background: #f1f5f9; cursor: pointer; font-size: 8px; line-height: 1;", disabled: !is_enabled, "▼" }
                                        }
                                    }
                                },
                                ControlType::PropertyGrid => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; overflow: auto; {base_field_bg} {style_back} {style_font};",
                                        div {
                                            style: "padding: 4px 6px; background: #e2e8f0; border-bottom: 1px solid #cbd5e1; font-size: 11px; font-weight: bold; color: #475569;",
                                            "Properties"
                                        }
                                        div {
                                            style: "padding: 8px; font-size: 11px; color: #64748b;",
                                            "[ PropertyGrid ]"
                                        }
                                    }
                                },
                                ControlType::Splitter => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; background: #e2e8f0; cursor: col-resize; border-left: 1px solid #cbd5e1; border-right: 1px solid #cbd5e1;",
                                    }
                                },
                                ControlType::DataGrid => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; overflow: auto; {base_field_bg} {style_back} {style_font};",
                                        table {
                                            style: "width: 100%; border-collapse: collapse; font-size: 12px;",
                                            tbody {
                                                tr {
                                                    td {
                                                        style: "padding: 6px 8px; border: 1px solid #e2e8f0; color: #64748b; text-align: center;",
                                                        "[ DataGrid ]"
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                ControlType::PrintPreviewControl => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; background: #f1f5f9; display: flex; align-items: center; justify-content: center; {style_font};",
                                        div {
                                            style: "text-align: center; color: #64748b;",
                                            div { style: "font-size: 32px; margin-bottom: 8px;", "🖨" }
                                            div { style: "font-size: 11px;", "Print Preview" }
                                        }
                                    }
                                },
                                ControlType::ToolStripSeparator => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;",
                                        div { style: "width: 1px; height: 80%; background: #cbd5e1;" }
                                    }
                                },
                                ControlType::ToolStripButton => rsx! {
                                    button {
                                        style: "width: 100%; height: 100%; border: 1px solid transparent; background: transparent; cursor: pointer; border-radius: 3px; {style_font} {style_fore}; padding: 2px 4px;",
                                        disabled: !is_enabled,
                                        onclick: move |_| { handle_event(name_clone.clone(), "Click".to_string(), None); },
                                        "{text}"
                                    }
                                },
                                ControlType::ToolStripLabel => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; display: flex; align-items: center; padding: 0 4px; {style_font} {style_fore} {style_back};",
                                        "{text}"
                                    }
                                },
                                ControlType::ToolStripComboBox => rsx! {
                                    select {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 3px; {base_field_bg} {style_back} {style_font} {style_fore}; padding: 0 2px;",
                                        disabled: !is_enabled,
                                    }
                                },
                                ControlType::ToolStripDropDownButton | ControlType::ToolStripSplitButton => rsx! {
                                    button {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; background: #f8fafc; cursor: pointer; border-radius: 3px; {style_font} {style_fore}; padding: 2px 4px; display: flex; align-items: center; gap: 2px;",
                                        disabled: !is_enabled,
                                        onclick: move |_| { handle_event(name_clone.clone(), "Click".to_string(), None); },
                                        span { "{text}" }
                                        span { style: "font-size: 8px;", "▼" }
                                    }
                                },
                                ControlType::ToolStripTextBox => rsx! {
                                    input {
                                        style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 3px; {base_field_bg} {style_back} {style_font} {style_fore}; padding: 0 4px;",
                                        disabled: !is_enabled,
                                        value: "{text}",
                                        oninput: move |evt| {
                                            if let Some(frm) = runtime_form.write().as_mut() {
                                                if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                    ctrl.set_text(evt.value());
                                                }
                                            }
                                            handle_event(name_clone.clone(), "TextChanged".to_string(), None);
                                        }
                                    }
                                },
                                ControlType::ToolStripProgressBar => {
                                    let ts_val = control.properties.get_int("Value").unwrap_or(0);
                                    let ts_max = control.properties.get_int("Maximum").unwrap_or(100);
                                    let ts_pct = if ts_max > 0 { (ts_val * 100) / ts_max } else { 0 };
                                    let ts_pct_str = format!("{}%", ts_pct);
                                    rsx! {
                                        div {
                                            style: "width: 100%; height: 100%; background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 3px; overflow: hidden;",
                                            div {
                                                style: "height: 100%; background: #2563eb; width: {ts_pct_str}; transition: width 0.3s;",
                                            }
                                        }
                                    }
                                },
                                ControlType::DataGridView => rsx! {
                                    {
                                        // Get dynamic grid data from properties (set by DataSourceChanged)
                                        let grid_columns = control.properties.get_string_array("__grid_columns")
                                            .cloned()
                                            .unwrap_or_default();
                                        let grid_row_strs = control.properties.get_string_array("__grid_rows")
                                            .cloned()
                                            .unwrap_or_default();
                                        let grid_rows: Vec<Vec<String>> = grid_row_strs.iter()
                                            .map(|s| s.split('\t').map(|c| c.to_string()).collect())
                                            .collect();
                                        let has_data = !grid_columns.is_empty();

                                        let dgv_selected_row = control.properties.get_int("CurrentRowIndex").unwrap_or(-1);
                                        rsx! {
                                            div {
                                                style: "width: 100%; height: 100%; border: 1px solid #999; background: #f0f0f0; padding: 1px; overflow: auto;",
                                                table {
                                                    style: "width: 100%; background: white; border-collapse: separate; border-spacing: 0; font-size: 12px;",
                                                    thead {
                                                        tr {
                                                            th { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 2px solid #999; padding: 4px 6px; width: 30px; text-align: center; font-weight: normal; color: #333;", "" }
                                                            if has_data {
                                                                for col in &grid_columns {
                                                                    th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px; text-align: left; font-weight: bold; color: #222; cursor: default; white-space: nowrap;", "{col}" }
                                                                }
                                                            } else {
                                                                th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px;", "Column1" }
                                                                th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px;", "Column2" }
                                                            }
                                                        }
                                                    }
                                                    tbody {
                                                        if has_data {
                                                            for (ri, row) in grid_rows.iter().enumerate() {
                                                                {
                                                                    let row_idx = ri as i32;
                                                                    let dgv_name = name_clone.clone();
                                                                    let row_bg = if row_idx == dgv_selected_row { "background: #cce5ff;" } else { "" };
                                                                    rsx! {
                                                                        tr {
                                                                            key: "{ri}",
                                                                            style: "{row_bg} cursor: pointer;",
                                                                            onclick: move |_| {
                                                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&dgv_name) {
                                                                                        ctrl.properties.set("CurrentRowIndex", row_idx);
                                                                                    }
                                                                                }
                                                                                handle_event(dgv_name.clone(), "CellClick".to_string(), None);
                                                                                handle_event(dgv_name.clone(), "RowEnter".to_string(), None);
                                                                                handle_event(dgv_name.clone(), "SelectionChanged".to_string(), None);
                                                                                handle_event(dgv_name.clone(), "CurrentCellChanged".to_string(), None);
                                                                            },
                                                                            td { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center; padding: 2px 4px; color: #333; width: 30px; height: 22px;", "{ri}" }
                                                                            for cell in row {
                                                                                td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; white-space: nowrap; height: 22px; {row_bg}", "{cell}" }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        } else {
                                                            tr {
                                                                td { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center; padding: 2px 4px; width: 30px; height: 22px;", "" }
                                                                td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; height: 22px;", "" }
                                                                td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; height: 22px;", "" }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                // Fallback for others
                                _ => rsx! {
                                    div {
                                        style: "width: 100%; height: 100%; border: 1px solid #e2e8f0; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #666; {style_back}",
                                        "{text}"
                                    }
                                }
                            }}
                        }
                    }
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

#[allow(dead_code)]
fn event_type_from_name(name: &str) -> Option<EventType> {
    match name.to_lowercase().as_str() {
        "click" => Some(EventType::Click),
        "dblclick" => Some(EventType::DblClick),
        "doubleclick" => Some(EventType::DoubleClick),
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
        "mouseclick" => Some(EventType::MouseClick),
        "mousedoubleclick" => Some(EventType::MouseDoubleClick),
        "mousedown" => Some(EventType::MouseDown),
        "mouseup" => Some(EventType::MouseUp),
        "mousemove" => Some(EventType::MouseMove),
        "mouseenter" => Some(EventType::MouseEnter),
        "mouseleave" => Some(EventType::MouseLeave),
        "mousewheel" => Some(EventType::MouseWheel),
        "gotfocus" => Some(EventType::GotFocus),
        "lostfocus" => Some(EventType::LostFocus),
        "enter" => Some(EventType::Enter),
        "leave" => Some(EventType::Leave),
        "validated" => Some(EventType::Validated),
        "validating" => Some(EventType::Validating),
        "resize" => Some(EventType::Resize),
        "paint" => Some(EventType::Paint),
        "formclosing" => Some(EventType::FormClosing),
        "formclosed" => Some(EventType::FormClosed),
        "shown" => Some(EventType::Shown),
        "activated" => Some(EventType::Activated),
        "deactivate" => Some(EventType::Deactivate),
        "tick" => Some(EventType::Tick),
        "elapsed" => Some(EventType::Elapsed),
        "scroll" => Some(EventType::Scroll),
        "selectedvaluechanged" => Some(EventType::SelectedValueChanged),
        "cellclick" => Some(EventType::CellClick),
        "celldoubleclick" => Some(EventType::CellDoubleClick),
        "cellvaluechanged" => Some(EventType::CellValueChanged),
        "selectionchanged" => Some(EventType::SelectionChanged),
        "linkclicked" => Some(EventType::LinkClicked),
        "datechanged" => Some(EventType::DateChanged),
        "dateselected" => Some(EventType::DateSelected),
        "itemclicked" => Some(EventType::ItemClicked),
        "buttonclick" => Some(EventType::ButtonClick),
        "splittermoved" => Some(EventType::SplitterMoved),
        "maskinputrejected" => Some(EventType::MaskInputRejected),
        _ => None,
    }
}

/// Map Dioxus keyboard Key to Windows Forms Virtual Key code (VK_*).
#[allow(dead_code)]
fn dioxus_key_to_vk(key: &dioxus::prelude::Key) -> i32 {
    use dioxus::prelude::Key;
    match key {
        Key::Backspace => 8,
        Key::Tab => 9,
        Key::Enter => 13,
        Key::Shift => 16,
        Key::Control => 17,
        Key::Alt => 18,
        Key::Pause => 19,
        Key::CapsLock => 20,
        Key::Escape => 27,
        Key::PageUp => 33,
        Key::PageDown => 34,
        Key::End => 35,
        Key::Home => 36,
        Key::ArrowLeft => 37,
        Key::ArrowUp => 38,
        Key::ArrowRight => 39,
        Key::ArrowDown => 40,
        Key::Insert => 45,
        Key::Delete => 46,
        Key::Character(c) => {
            let ch = c.chars().next().unwrap_or('\0');
            match ch {
                '0'..='9' => ch as i32,              // 0x30–0x39
                'a'..='z' => (ch as i32) - 32,       // VK uses uppercase: 0x41–0x5A
                'A'..='Z' => ch as i32,
                ' ' => 32,
                '+' | '=' => 187,  // VK_OEM_PLUS
                '-' | '_' => 189,  // VK_OEM_MINUS
                ',' | '<' => 188,  // VK_OEM_COMMA
                '.' | '>' => 190,  // VK_OEM_PERIOD
                '/' | '?' => 191,  // VK_OEM_2
                ';' | ':' => 186,  // VK_OEM_1
                '\'' | '"' => 222, // VK_OEM_7
                '[' | '{' => 219,  // VK_OEM_4
                ']' | '}' => 221,  // VK_OEM_6
                '\\' | '|' => 220, // VK_OEM_5
                '`' | '~' => 192, // VK_OEM_3
                _ => ch as i32,
            }
        }
        Key::F1 => 112,
        Key::F2 => 113,
        Key::F3 => 114,
        Key::F4 => 115,
        Key::F5 => 116,
        Key::F6 => 117,
        Key::F7 => 118,
        Key::F8 => 119,
        Key::F9 => 120,
        Key::F10 => 121,
        Key::F11 => 122,
        Key::F12 => 123,
        Key::NumLock => 144,
        Key::ScrollLock => 145,
        _ => 0,
    }
}

// ---------------------------------------------------------------------------
// Control Sync Helpers
//
// The Form struct (vybe_forms) is the UI source of truth — the renderer reads
// from it.  The interpreter's form instance object is the code source of
// truth — VB handlers read/write to it via `Me.ctrl.Prop`.  These two helpers
// synchronise the two representations cleanly:
//
//   sync_ui_to_instance    – before event dispatch (push UI state into interp)
//   sync_instance_to_ui    – after  event dispatch (pull code changes into UI)
// ---------------------------------------------------------------------------

/// Push current UI state (text-box values, checkbox states …) from the Form
/// struct into the interpreter's form instance object so VB code sees the
/// latest values.
fn sync_ui_to_instance(
    form: &Form,
    form_obj: &Rc<RefCell<vybe_runtime::value::ObjectData>>,
) {
    let form_ref = form_obj.borrow();
    for ctrl in &form.controls {
        let key = ctrl.name.to_lowercase();
        if let Some(Value::Object(ctrl_obj)) = form_ref.fields.get(&key) {
            let mut ctrl_ref = ctrl_obj.borrow_mut();
            // Text  (textbox input, label changes, …)
            if let Some(text) = ctrl.get_text() {
                ctrl_ref.fields.insert("text".to_string(), Value::String(text.to_string()));
            }
            // Enabled
            ctrl_ref.fields.insert("enabled".to_string(), Value::Boolean(ctrl.is_enabled()));
            // Visible
            ctrl_ref.fields.insert("visible".to_string(), Value::Boolean(ctrl.is_visible()));
            // Checked
            if let Some(ck) = ctrl.properties.get_bool("Checked") {
                ctrl_ref.fields.insert("checked".to_string(), Value::Boolean(ck));
            }
            // Value (trackbar, numeric up-down, …)
            if let Some(v) = ctrl.properties.get_int("Value") {
                ctrl_ref.fields.insert("value".to_string(), Value::Integer(v));
            }
            // SelectedIndex
            if let Some(si) = ctrl.properties.get_int("SelectedIndex") {
                ctrl_ref.fields.insert("selectedindex".to_string(), Value::Integer(si));
            }
        }
    }
}

/// Pull property changes made by VB code from the interpreter's form
/// instance object back into the Form struct so the renderer shows them.
fn sync_instance_to_ui(
    form: &mut Form,
    form_obj: &Rc<RefCell<vybe_runtime::value::ObjectData>>,
) {
    let form_ref = form_obj.borrow();

    // Sync form-level Text/Caption
    if let Some(Value::String(s)) = form_ref.fields.get("text") {
        form.text = s.clone();
    }

    for ctrl in form.controls.iter_mut() {
        let key = ctrl.name.to_lowercase();
        if let Some(Value::Object(ctrl_obj)) = form_ref.fields.get(&key) {
            let ctrl_fields = ctrl_obj.borrow();
            // Text
            if let Some(Value::String(s)) = ctrl_fields.fields.get("text") {
                ctrl.set_text(s.clone());
            }
            // Multiline
            if let Some(Value::Boolean(b)) = ctrl_fields.fields.get("multiline") {
                ctrl.properties.set("Multiline", *b);
            }
            // ReadOnly
            if let Some(Value::Boolean(b)) = ctrl_fields.fields.get("readonly") {
                ctrl.properties.set("ReadOnly", *b);
            }
            // Enabled
            if let Some(val) = ctrl_fields.fields.get("enabled") {
                let en = match val {
                    Value::Boolean(b) => *b,
                    Value::Integer(i) => *i != 0,
                    _ => true,
                };
                ctrl.properties.set_raw("Enabled", vybe_forms::PropertyValue::Boolean(en));
            }
            // Visible
            if let Some(val) = ctrl_fields.fields.get("visible") {
                let vis = match val {
                    Value::Boolean(b) => *b,
                    Value::Integer(i) => *i != 0,
                    _ => true,
                };
                ctrl.properties.set_raw("Visible", vybe_forms::PropertyValue::Boolean(vis));
            }
            // BackColor
            if let Some(val) = ctrl_fields.fields.get("backcolor") {
                if let Some(css) = value_to_css_color(val) {
                    ctrl.set_back_color(css);
                }
            }
            // ForeColor
            if let Some(val) = ctrl_fields.fields.get("forecolor") {
                if let Some(css) = value_to_css_color(val) {
                    ctrl.set_fore_color(css);
                }
            }
            // Checked
            if let Some(Value::Boolean(b)) = ctrl_fields.fields.get("checked") {
                ctrl.properties.set_raw("Checked", vybe_forms::PropertyValue::Boolean(*b));
            }
            // SelectedIndex
            if let Some(Value::Integer(i)) = ctrl_fields.fields.get("selectedindex") {
                ctrl.properties.set_raw("SelectedIndex", vybe_forms::PropertyValue::Integer(*i));
            }
            // Value
            if let Some(Value::Integer(i)) = ctrl_fields.fields.get("value") {
                ctrl.properties.set_raw("Value", vybe_forms::PropertyValue::Integer(*i));
            }
            // WebBrowser URL/HTML
            if let Some(Value::String(s)) = ctrl_fields.fields.get("url") {
                ctrl.properties.set("URL", s.clone());
            }
            if let Some(Value::String(s)) = ctrl_fields.fields.get("html") {
                ctrl.properties.set("HTML", s.clone());
            } else if let Some(Value::String(s)) = ctrl_fields.fields.get("documenttext") {
                ctrl.properties.set("HTML", s.clone());
            }
        }
    }
}

/// After the interpreter creates the form instance (via Sub New /
/// InitializeComponent), some controls may not have been created properly
/// (e.g. the interpreter couldn't evaluate all designer expressions).
/// This ensures every control from the designer-parsed Form has a
/// corresponding object field on the instance with correct values.
fn ensure_controls_on_instance(
    form: &Form,
    form_obj: &Rc<RefCell<vybe_runtime::value::ObjectData>>,
) {
    let mut form_ref = form_obj.borrow_mut();
    for ctrl in &form.controls {
        let key = ctrl.name.to_lowercase();
        let existing = form_ref.fields.get(&key);

        // If the field is missing, Nothing, or still a plain String placeholder,
        // create a proper control object from the designer data.
        let needs_create = match existing {
            None | Some(Value::Nothing) | Some(Value::String(_)) => true,
            _ => false,
        };

        if needs_create {
            let ctrl_obj = build_control_object(ctrl);
            form_ref.fields.insert(key, ctrl_obj);
        } else if let Some(Value::Object(ctrl_obj)) = existing {
            // The object exists (created by InitializeComponent).
            // Fill in any missing properties from the designer data
            // but do NOT overwrite properties that InitializeComponent set.
            let mut cr = ctrl_obj.borrow_mut();
            if !cr.fields.contains_key("name") {
                cr.fields.insert("name".to_string(), Value::String(ctrl.name.clone()));
            }
            if !cr.fields.contains_key("__is_control") {
                cr.fields.insert("__is_control".to_string(), Value::Boolean(true));
            }
            // If InitializeComponent created the control but text is empty while
            // designer has a value, prefer the designer value. This handles the
            // case where InitializeComponent did `Me.btn = New Button()` but the
            // interpreter skipped the `Me.btn.Text = "1"` line.
            let has_empty_text = cr.fields.get("text")
                .map(|v| matches!(v, Value::String(s) if s.is_empty()))
                .unwrap_or(true);
            if has_empty_text {
                if let Some(text) = ctrl.get_text() {
                    if !text.is_empty() {
                        cr.fields.insert("text".to_string(), Value::String(text.to_string()));
                    }
                }
            }
        }
    }
}

/// Build a Value::Object representing a single control from its designer data.
fn build_control_object(ctrl: &vybe_forms::Control) -> Value {
    let mut fields = HashMap::new();
    fields.insert("__type".to_string(), Value::String(ctrl.control_type.as_str().to_string()));
    fields.insert("__is_control".to_string(), Value::Boolean(true));
    fields.insert("name".to_string(), Value::String(ctrl.name.clone()));
    fields.insert("text".to_string(), Value::String(
        ctrl.get_text().unwrap_or("").to_string()
    ));
    fields.insert("enabled".to_string(), Value::Boolean(ctrl.is_enabled()));
    fields.insert("visible".to_string(), Value::Boolean(ctrl.is_visible()));
    fields.insert("left".to_string(), Value::Integer(ctrl.bounds.x));
    fields.insert("top".to_string(), Value::Integer(ctrl.bounds.y));
    fields.insert("width".to_string(), Value::Integer(ctrl.bounds.width));
    fields.insert("height".to_string(), Value::Integer(ctrl.bounds.height));
    fields.insert("tag".to_string(), Value::Nothing);
    fields.insert("tabindex".to_string(), Value::Integer(ctrl.tab_index as i32));
    if let Some(bc) = ctrl.get_back_color() {
        fields.insert("backcolor".to_string(), Value::String(bc.to_string()));
    }
    if let Some(fc) = ctrl.get_fore_color() {
        fields.insert("forecolor".to_string(), Value::String(fc.to_string()));
    }
    if let Some(fnt) = ctrl.get_font() {
        fields.insert("font".to_string(), Value::String(fnt.to_string()));
    }
    if let Some(ck) = ctrl.properties.get_bool("Checked") {
        fields.insert("checked".to_string(), Value::Boolean(ck));
    }
    if let Some(v) = ctrl.properties.get_int("Value") {
        fields.insert("value".to_string(), Value::Integer(v));
    }
    if let Some(si) = ctrl.properties.get_int("SelectedIndex") {
        fields.insert("selectedindex".to_string(), Value::Integer(si));
    }

    Value::Object(Rc::new(RefCell::new(ObjectData {
        class_name: ctrl.control_type.as_str().to_string(),
        fields,
        drawing_commands: Vec::new(),
    })))
}

/// Convert a runtime Value to a CSS color string.
/// Handles Value::String (pass-through), Value::Object (Color with r/g/b fields).
fn value_to_css_color(value: &vybe_runtime::Value) -> Option<String> {
    use vybe_runtime::Value;
    match value {
        Value::String(s) if !s.is_empty() => Some(s.clone()),
        Value::Object(obj_ref) => {
            let b = obj_ref.borrow();
            let extract_u32 = |v: &Value| -> Option<u32> {
                match v {
                    Value::Byte(b) => Some(*b as u32),
                    Value::Integer(i) => Some(*i as u32),
                    _ => None,
                }
            };
            let r = b.fields.get("r").and_then(extract_u32);
            let g = b.fields.get("g").and_then(extract_u32);
            let b_val = b.fields.get("b").and_then(extract_u32);
            if let (Some(r), Some(g), Some(b)) = (r, g, b_val) {
                Some(format!("#{:02X}{:02X}{:02X}", r, g, b))
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Convert a runtime Value to a CSS font string.
/// Handles Value::String (pass-through), Value::Object (Font with name/size fields).
fn value_to_css_font(value: &vybe_runtime::Value) -> Option<String> {
    use vybe_runtime::Value;
    match value {
        Value::String(s) if !s.is_empty() => Some(s.clone()),
        Value::Object(obj_ref) => {
            let b = obj_ref.borrow();
            let name = b.fields.get("name")
                .or_else(|| b.fields.get("fontfamily"))
                .and_then(|v| if let Value::String(s) = v { Some(s.clone()) } else { None })
                .unwrap_or_else(|| "Segoe UI".to_string());
            let size = b.fields.get("size")
                .or_else(|| b.fields.get("sizeininpoints"))
                .and_then(|v| match v {
                    Value::Double(d) => Some(*d),
                    Value::Single(f) => Some(*f as f64),
                    Value::Integer(i) => Some(*i as f64),
                    _ => None,
                })
                .unwrap_or(9.0);
            let bold = b.fields.get("bold").and_then(|v| v.as_bool().ok()).unwrap_or(false);
            let italic = b.fields.get("italic").and_then(|v| v.as_bool().ok()).unwrap_or(false);
            let mut parts = Vec::new();
            if italic { parts.push("italic".to_string()); }
            if bold { parts.push("bold".to_string()); }
            parts.push(format!("{}px", size));
            parts.push(format!("'{}'", name));
            Some(parts.join(" "))
        }
        _ => None,
    }
}

fn process_side_effects(
    interp: &mut Interpreter,
    rp: RuntimeProject,
    runtime_form: &mut Signal<Option<Form>>,
    msgbox_content: &mut Signal<Option<String>>,
    wb_html: &mut Signal<std::collections::HashMap<String, String>>,
) {
    while let Some(effect) = interp.side_effects.pop_front() {
        match effect {
            RuntimeSideEffect::MsgBox(msg) => {
                msgbox_content.set(Some(msg));
            }
            RuntimeSideEffect::Repaint { control_name: _ } => {
                // TODO: Implement actual UI repaint triggering
            }
            RuntimeSideEffect::RunApplication { form_name } => {
                let project_read = rp.project.read();
                if let Some(proj) = project_read.as_ref() {
                     if let Some(form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&form_name)) {
                         // Found the form definition
                         let mut new_form = form_module.form.clone();
                         
                         // Sync from the runtime instance
                         if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                             sync_instance_to_ui(&mut new_form, &form_obj);
                             
                             // Trigger Load event
                             let load_args = interp.make_event_handler_args(&form_name, "Load");
                             // Note: get_event_handlers needs form class name (form_name) and control name "Me"
                             let handlers = interp.get_event_handlers(&form_name.to_lowercase(), "Me", "Load");
                             for handler in handlers {
                                 let _ = interp.call_method_on_object(&form_obj, &handler, &load_args);
                             }
                             
                             // Trigger Shown event
                             let shown_args = interp.make_event_handler_args(&form_name, "Shown");
                             let handlers = interp.get_event_handlers(&form_name.to_lowercase(), "Me", "Shown");
                             for handler in handlers {
                                 let _ = interp.call_method_on_object(&form_obj, &handler, &shown_args);
                             }
                         }
                         
                         runtime_form.set(Some(new_form));
                     }
                }
            }
            RuntimeSideEffect::PropertyChange { object, property, value } => {
                let mut switched = false;

                let project_read = rp.project.read();
                if let Some(proj) = project_read.as_ref() {
                    if let Some(other_form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&object)) {
                        let current_is_it = if let Some(cf) = &*runtime_form.peek() {
                            cf.name.eq_ignore_ascii_case(&object)
                        } else {
                            false
                        };

                        if !current_is_it {
                            if property.eq_ignore_ascii_case("Visible")
                                && (value.as_bool().unwrap_or(false) || value.as_string() == "True")
                            {
                                let switch_code = if other_form_module.is_vbnet() {
                                    format!("{}\n{}", other_form_module.get_designer_code(), other_form_module.get_user_code())
                                } else {
                                    other_form_module.get_user_code().to_string()
                                };
                                match parse_program(&switch_code) {
                                    Ok(prog) => {
                                        if let Err(e) = interp.load_module(&other_form_module.form.name, &prog) {
                                            println!("Error loading new form code: {:?}", e);
                                        } else {
                                            let load_args = interp.make_event_handler_args(&other_form_module.form.name, "Load");
                            let _ = interp.call_event_handler(&format!("{}_Load", other_form_module.form.name), &load_args);
                                            runtime_form.set(Some(other_form_module.form.clone()));
                                            switched = true;
                                        }
                                    }
                                    Err(_) => {
                                        println!("Failed to parse new form code");
                                    }
                                }
                            }
                        }
                    }
                }

                if !switched {
                    if let Some(frm) = runtime_form.write().as_mut() {
                        let (form_part, control_part) = if let Some(idx) = object.find('.') {
                            (&object[..idx], &object[idx + 1..])
                        } else {
                            ("", object.as_str())
                        };

                        let is_for_this_form = form_part.is_empty() || form_part.eq_ignore_ascii_case(&frm.name);

                        if is_for_this_form {
                            if control_part.eq_ignore_ascii_case(&frm.name)
                                || (form_part.is_empty() && object.eq_ignore_ascii_case(&frm.name))
                            {
                                if property.eq_ignore_ascii_case("Caption") || property.eq_ignore_ascii_case("Text") {
                                    frm.text = value.as_string();
                                }
                            } else {
                                if let Some(ctrl) = frm.get_control_by_name_mut(control_part) {
                                    match property.to_lowercase().as_str() {
                                        "text" => {
                                            let text_val = value.as_string();
                                            ctrl.set_text(text_val.clone());
                                            // Mirror the new text back into form_obj so
                                            // needs_sync doesn't detect a stale mismatch
                                            // and overwrite this binding-driven update.
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("text".to_string(), Value::String(text_val.clone()));
                                                }
                                            }
                                            if ctrl.control_type == vybe_forms::ControlType::ListBox {
                                                let items: Vec<String> = text_val
                                                    .split('|')
                                                    .map(|s| s.trim().to_string())
                                                    .filter(|s| !s.is_empty())
                                                    .collect();
                                                if !items.is_empty() {
                                                    ctrl.set_list_items(items);
                                                }
                                            }
                                        }
                                        "caption" => {
                                            // VB6 compat: Caption maps to Text in .NET
                                            let cap_val = value.as_string();
                                            ctrl.set_text(cap_val.clone());
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("text".to_string(), Value::String(cap_val));
                                                }
                                            }
                                        }
                                        "left" => { if let vybe_runtime::Value::Integer(v) = &value { ctrl.bounds.x = *v; } },
                                        "top" => { if let vybe_runtime::Value::Integer(v) = &value { ctrl.bounds.y = *v; } },
                                        "width" => { if let vybe_runtime::Value::Integer(v) = &value { ctrl.bounds.width = *v; } },
                                        "height" => { if let vybe_runtime::Value::Integer(v) = &value { ctrl.bounds.height = *v; } },
                                        "visible" => { ctrl.properties.set_raw("Visible", vybe_forms::PropertyValue::Boolean(value.as_bool().unwrap_or(true))); },
                                        "enabled" => { ctrl.properties.set_raw("Enabled", vybe_forms::PropertyValue::Boolean(value.as_bool().unwrap_or(true))); },
                                        "backcolor" => {
                                            if let Some(css) = value_to_css_color(&value) {
                                                ctrl.set_back_color(css);
                                            }
                                        },
                                        "forecolor" => {
                                            if let Some(css) = value_to_css_color(&value) {
                                                ctrl.set_fore_color(css);
                                            }
                                        },
                                        "font" => {
                                            if let Some(css) = value_to_css_font(&value) {
                                                ctrl.set_font(css);
                                            }
                                        },
                                        // CheckBox / RadioButton
                                        "checked" => {
                                            let checked = match &value {
                                                vybe_runtime::Value::Boolean(b) => *b,
                                                vybe_runtime::Value::Integer(i) => *i != 0,
                                                vybe_runtime::Value::String(s) => s == "1" || s.eq_ignore_ascii_case("true"),
                                                _ => false,
                                            };
                                            ctrl.properties.set_raw("Checked", vybe_forms::PropertyValue::Boolean(checked));
                                            use vybe_forms::properties::PropertyValue as PV;
                                            ctrl.properties.set_raw("Value", PV::Integer(if checked { 1 } else { 0 }));
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("checked".to_string(), Value::Boolean(checked));
                                                }
                                            }
                                        }
                                        // NumericUpDown, TrackBar, Scrollbars
                                        "value" => {
                                            let ival = match &value {
                                                vybe_runtime::Value::Integer(i) => *i,
                                                vybe_runtime::Value::Double(d) => *d as i32,
                                                vybe_runtime::Value::String(s) => s.parse().unwrap_or(0),
                                                _ => 0,
                                            };
                                            ctrl.properties.set_raw("Value", vybe_forms::PropertyValue::Integer(ival));
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("value".to_string(), Value::Integer(ival));
                                                }
                                            }
                                        }
                                        // ComboBox / ListBox selected index
                                        "selectedindex" => {
                                            let idx = match &value {
                                                vybe_runtime::Value::Integer(i) => *i,
                                                vybe_runtime::Value::String(s) => s.parse().unwrap_or(-1),
                                                _ => -1,
                                            };
                                            ctrl.properties.set_raw("SelectedIndex", vybe_forms::PropertyValue::Integer(idx));
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("selectedindex".to_string(), Value::Integer(idx));
                                                }
                                            }
                                        }
                                        // ComboBox / ListBox selected value (map to text)
                                        "selectedvalue" => {
                                            let sv = value.as_string();
                                            ctrl.set_text(sv.clone());
                                            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                                let form_borrow = form_obj.borrow();
                                                if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&control_part.to_lowercase()) {
                                                    ctrl_obj.borrow_mut().fields.insert("text".to_string(), Value::String(sv));
                                                }
                                            }
                                        }
                                        "url" => {
                                            ctrl.properties.set("URL", value.as_string());
                                            let url = value.as_string();
                                            let ctrl_id = control_part.to_lowercase();
                                            // Clear HTML content on Navigate
                                            wb_html.write().remove(&ctrl_id);
                                            // Only fetch if URL is new or changed
                                            let mut do_fetch = true;
                                            if let Some(existing) = wb_html.read().get(&ctrl_id) {
                                                if let Some(v) = ctrl.properties.get("URL") {
                                                    if let vybe_forms::PropertyValue::String(s) = v {
                                                        if s == &url && !existing.is_empty() {
                                                            do_fetch = false;
                                                        }
                                                    }
                                                }
                                            }
                                            if url.is_empty() || url == "about:blank" {
                                                wb_html.write().remove(&ctrl_id);
                                            } else if do_fetch {
                                                // Fetch HTML on the Rust side, store in the wb_html signal.
                                                eprintln!("[WebBrowser] fetching URL: {}", url);
                                                let output = std::process::Command::new("curl")
                                                    .args(&["-s", "-L", "-k", "--max-time", "10", &url])
                                                    .output();
                                                let html = match output {
                                                    Ok(out) if out.status.success() => {
                                                        let raw = String::from_utf8_lossy(&out.stdout).to_string();
                                                        let base_tag = format!("<base href=\"{}\" target=\"_self\">", url);
                                                        let nav_script = format!(
                                                            "<script>\
                                                             document.addEventListener('click',function(e){{\
                                                               var a=e.target.closest('a');\
                                                               if(a&&a.href){{\
                                                                 e.preventDefault();\
                                                                 e.stopPropagation();\
                                                                 window.__vybe_nav=a.href;\
                                                               }}\
                                                             }},true);\
                                                             </script>"
                                                        );
                                                        if let Some(pos) = raw.to_lowercase().find("<head") {
                                                            if let Some(end) = raw[pos..].find('>') {
                                                                format!("{}{}{}{}", &raw[..pos + end + 1], base_tag, &raw[pos + end + 1..], nav_script)
                                                            } else {
                                                                format!("{}{}{}", base_tag, raw, nav_script)
                                                            }
                                                        } else {
                                                            format!("{}{}{}", base_tag, raw, nav_script)
                                                        }
                                                    }
                                                    _ => format!("<html><body><p style='color:red'>Failed to load: {}</p></body></html>", url),
                                                };
                                                eprintln!("[WebBrowser] fetched {} bytes", html.len());
                                                wb_html.write().insert(ctrl_id, html);
                                            }
                                        }
                                        "html" | "documenttext" => {
                                            ctrl.properties.set("HTML", value.as_string());
                                            let html = value.as_string();
                                            let ctrl_id = control_part.to_lowercase();
                                            wb_html.write().insert(ctrl_id, html);
                                        }
                                        _ => {
                                            let prop_val = match &value {
                                                vybe_runtime::Value::Integer(i) => Some(vybe_forms::PropertyValue::Integer(*i)),
                                                vybe_runtime::Value::String(s) => Some(vybe_forms::PropertyValue::String(s.clone())),
                                                vybe_runtime::Value::Boolean(b) => Some(vybe_forms::PropertyValue::Boolean(*b)),
                                                vybe_runtime::Value::Double(d) => Some(vybe_forms::PropertyValue::Double(*d)),
                                                // For Object values, extract the "name" field as a string identifier
                                                // (e.g. BindingSource property set to a BindingSource object → store its name)
                                                vybe_runtime::Value::Object(obj_ref) => {
                                                    obj_ref.borrow().fields.get("name")
                                                        .and_then(|v| if let vybe_runtime::Value::String(s) = v {
                                                            if !s.is_empty() { Some(vybe_forms::PropertyValue::String(s.clone())) } else { None }
                                                        } else { None })
                                                }
                                                // Skip Nothing, Array, etc.
                                                _ => None,
                                            };
                                            if let Some(pv) = prop_val {
                                                ctrl.properties.set_raw(&property, pv);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            RuntimeSideEffect::ConsoleOutput(msg) => {
                print!("{}", msg);
            }
            RuntimeSideEffect::ConsoleClear => {
                // Console cleared
            }
            RuntimeSideEffect::InputBox { .. } => {}
    
            
            RuntimeSideEffect::DataSourceChanged { control_name, columns, rows } => {
                // Update the DataGridView control's grid data
                if let Some(frm) = runtime_form.write().as_mut() {
                    if let Some(ctrl) = frm.get_control_by_name_mut(&control_name) {
                        // Store columns and rows as serialized JSON in properties
                        ctrl.properties.set_raw("__grid_columns",
                            vybe_forms::PropertyValue::StringArray(columns.clone()));
                        let row_strs: Vec<String> = rows.iter()
                            .map(|r| r.join("\t"))
                            .collect();
                        ctrl.properties.set_raw("__grid_rows",
                            vybe_forms::PropertyValue::StringArray(row_strs));
                    }
                }
            }
            RuntimeSideEffect::BindingPositionChanged { binding_source_name, position, count } => {
                // Update BindingNavigator display for navigators linked to this BindingSource
                let mut updated_navs: Vec<(String, String)> = Vec::new();
                if let Some(frm) = runtime_form.write().as_mut() {
                    for ctrl in &mut frm.controls {
                        if matches!(ctrl.control_type, vybe_forms::ControlType::BindingNavigator) {
                            let ctrl_bs = ctrl.properties.get_string("BindingSource").unwrap_or_default();
                            if ctrl_bs.eq_ignore_ascii_case(&binding_source_name) {
                                let count_text = format!("{} of {}", position + 1, count);
                                ctrl.set_text(count_text.clone());
                                updated_navs.push((ctrl.name.clone(), count_text));
                            }
                        }
                    }
                }
                // Mirror nav text back into form_obj so needs_sync doesn't overwrite it
                if !updated_navs.is_empty() {
                    if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                        let form_borrow = form_obj.borrow();
                        for (nav_name, nav_text) in updated_navs {
                            if let Some(Value::Object(ctrl_obj)) = form_borrow.fields.get(&nav_name.to_lowercase()) {
                                ctrl_obj.borrow_mut().fields.insert("text".to_string(), Value::String(nav_text));
                            }
                        }
                    }
                }
            }
            RuntimeSideEffect::FormClose { form_name } => {
                // Fire FormClosing event, check Cancel, then fire FormClosed, then hide
                let closing_args = interp.make_event_handler_args(&form_name, "FormClosing");
                let form_name_lower = form_name.to_lowercase();
                // Dispatch using get_event_handlers (Handles clause + AddHandler) first,
                // then fall back to conventional Form1_FormClosing naming.
                let closing_handlers = interp.get_event_handlers(&form_name_lower, "Me", "FormClosing");
                if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                    if !closing_handlers.is_empty() {
                        for handler in &closing_handlers {
                            let _ = interp.call_method_on_object(&form_obj, handler, &closing_args);
                        }
                    } else {
                        let _ = interp.call_method_on_object(&form_obj, &format!("{}_FormClosing", form_name), &closing_args);
                    }
                } else {
                    let _ = interp.call_event_handler(&format!("{}_FormClosing", form_name), &closing_args);
                }
                // Check if Cancel was set to True on the EventArgs
                let cancel = if let Value::Object(ref ea) = closing_args[1] {
                    ea.borrow().fields.get("cancel").map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false)
                } else {
                    false
                };
                if !cancel {
                    let closed_args = interp.make_event_handler_args(&form_name, "FormClosed");
                    let closed_handlers = interp.get_event_handlers(&form_name_lower, "Me", "FormClosed");
                    if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                        if !closed_handlers.is_empty() {
                            for handler in &closed_handlers {
                                let _ = interp.call_method_on_object(&form_obj, handler, &closed_args);
                            }
                        } else {
                            let _ = interp.call_method_on_object(&form_obj, &format!("{}_FormClosed", form_name), &closed_args);
                        }
                    } else {
                        let _ = interp.call_event_handler(&format!("{}_FormClosed", form_name), &closed_args);
                    }
                    // Hide the form — read first, then set
                    let should_close = runtime_form.read().as_ref()
                        .map(|frm| frm.name.eq_ignore_ascii_case(&form_name))
                        .unwrap_or(false);
                    if should_close {
                        runtime_form.set(None);
                    }
                }
            }
            RuntimeSideEffect::FormShowDialog { form_name } => {
                // Show another form as modal (same as Show for now — full modal requires blocking)
                let project_read = rp.project.read();
                if let Some(proj) = project_read.as_ref() {
                    if let Some(form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&form_name)) {
                        let switch_code = if form_module.is_vbnet() {
                            format!("{}\n{}", form_module.get_designer_code(), form_module.get_user_code())
                        } else {
                            form_module.get_user_code().to_string()
                        };
                        if let Ok(prog) = parse_program(&switch_code) {
                            let _ = interp.load_module(&form_module.form.name, &prog);
                            let load_args = interp.make_event_handler_args(&form_module.form.name, "Load");
                            let _ = interp.call_event_handler(&format!("{}_Load", form_module.form.name), &load_args);
                            runtime_form.set(Some(form_module.form.clone()));
                        }
                    }
                }
            }
            RuntimeSideEffect::AddControl { form_name, control_name, control_type, left, top, width, height, parent_name } => {
                // Dynamically add a control at runtime
                if let Some(frm) = runtime_form.write().as_mut() {
                    if frm.name.eq_ignore_ascii_case(&form_name) || form_name.is_empty() {
                        // If a control with this name already exists (from design-time),
                        // update its parent_id instead of adding a duplicate.
                        let existing = frm.controls.iter().position(|c| c.name.eq_ignore_ascii_case(&control_name));
                        if let Some(idx) = existing {
                            if !parent_name.is_empty() {
                                let parent_id = frm.controls.iter()
                                    .find(|c| c.name.eq_ignore_ascii_case(&parent_name))
                                    .map(|c| c.id);
                                if let Some(pid) = parent_id {
                                    frm.controls[idx].parent_id = Some(pid);
                                }
                            } else {
                                frm.controls[idx].parent_id = None;
                            }
                        } else {
                            let ct = vybe_forms::ControlType::from_name(&control_type);
                            if let Some(ct) = ct {
                                let mut ctrl = vybe_forms::Control::new(ct, control_name.clone(), left, top);
                                ctrl.bounds.width = width;
                                ctrl.bounds.height = height;
                                // Resolve parent control by name
                                if !parent_name.is_empty() {
                                    if let Some(parent) = frm.controls.iter().find(|c| c.name.eq_ignore_ascii_case(&parent_name)) {
                                        ctrl.parent_id = Some(parent.id);
                                    }
                                }
                                frm.controls.push(ctrl);
                            }
                        }
                    }
                }
            }
        }
    }

    // ── Sync interpreter → UI after side effects ────────────────────────
    // The form instance key is stored as `__form_instance__` in the
    // interpreter's global env.  Read from it and push changes into the
    // Form struct that drives the renderer.
    if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
        let needs_sync = if let Some(frm) = runtime_form.peek().as_ref() {
            let form_borrow = form_obj.borrow();
            let mut changed = false;
            for control in &frm.controls {
                if let Some(Value::Object(ctrl_obj)) =
                    form_borrow.fields.get(&control.name.to_lowercase())
                {
                    let ctrl_fields = ctrl_obj.borrow();
                    if let Some(Value::String(s)) = ctrl_fields.fields.get("text") {
                        let current = control.get_text().map(|t| t.to_string()).unwrap_or_default();
                        if *s != current { changed = true; break; }
                    }
                    if let Some(val) = ctrl_fields.fields.get("enabled") {
                        let new_en = match val {
                            Value::Boolean(b) => *b,
                            Value::Integer(i) => *i != 0,
                            _ => true,
                        };
                        if new_en != control.is_enabled() { changed = true; break; }
                    }
                    if let Some(val) = ctrl_fields.fields.get("visible") {
                        let new_vis = match val {
                            Value::Boolean(b) => *b,
                            Value::Integer(i) => *i != 0,
                            _ => true,
                        };
                        if new_vis != control.is_visible() { changed = true; break; }
                    }
                }
            }
            changed
        } else {
            false
        };

        if needs_sync {
            if let Some(frm) = runtime_form.write().as_mut() {
                sync_instance_to_ui(frm, &form_obj);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// FormRunner – the single, shared Dioxus component used by both the editor
// and the standalone CLI shell.
// ---------------------------------------------------------------------------

#[component]
pub fn FormRunner() -> Element {
    let mut rp = use_context::<RuntimeProject>();

    let mut interpreter = use_signal(|| None::<Interpreter>);
    let mut runtime_form = use_signal(|| None::<Form>);
    let mut msgbox_content = use_signal(|| None::<String>);
    let mut parse_error = use_signal(|| None::<String>);
    let mut handling_event = use_signal(|| false);

    // ── Console mode state ──────────────────────────────────────────────
    let mut console_output = use_signal(String::new);
    let mut console_waiting_input = use_signal(|| false);
    let mut console_finished = use_signal(|| false);
    let mut console_input_line = use_signal(String::new);
    let mut is_console_mode = use_signal(|| false);

    // Channel endpoints wrapped so they can live in signals.
    type TxWrap = Rc<RefCell<Option<mpsc::Sender<String>>>>;
    type RxWrap = Rc<RefCell<Option<mpsc::Receiver<ConsoleMessage>>>>;
    let mut console_input_tx: Signal<Option<TxWrap>> = use_signal(|| None);
    let mut console_rx: Signal<Option<RxWrap>> = use_signal(|| None);

    // WebBrowser HTML content – stored separately so form-sync can't wipe it.
    let mut wb_html: Signal<std::collections::HashMap<String, String>> =
        use_signal(|| std::collections::HashMap::new());

    // ── Poll for link clicks inside WebBrowser content ───────────────
    // The injected script sets window.__vybe_nav when a link is clicked.
    // We poll every 200ms, fetch the URL via curl, and update the signal.
    {
        let mut wb_html_poll = wb_html.clone();
        use_future(move || async move {
            loop {
                tokio::time::sleep(std::time::Duration::from_millis(200)).await;
                let mut eval = document::eval(
                    "if(window.__vybe_nav){var u=window.__vybe_nav;window.__vybe_nav=null;dioxus.send(u);}else{dioxus.send('');}"
                );
                if let Ok(url) = eval.recv::<String>().await {
                    if !url.is_empty() {
                        eprintln!("[WebBrowser] link click: {}", url);
                        // Fetch the linked page
                        let output = std::process::Command::new("curl")
                            .args(&["-s", "-L", "-k", "--max-time", "10", &url])
                            .output();

                        let html = match output {
                            Ok(out) if out.status.success() => {
                                let raw = String::from_utf8_lossy(&out.stdout).to_string();
                                let base_tag = format!("<base href=\"{}\" target=\"_self\">", url);
                                let nav_script = "<script>\
                                    document.addEventListener('click',function(e){\
                                      var a=e.target.closest('a');\
                                      if(a&&a.href){\
                                        e.preventDefault();\
                                        e.stopPropagation();\
                                        window.__vybe_nav=a.href;\
                                      }\
                                    },true);\
                                    </script>";
                                if let Some(pos) = raw.to_lowercase().find("<head") {
                                    if let Some(end) = raw[pos..].find('>') {
                                        format!("{}{}{}{}", &raw[..pos + end + 1], base_tag, &raw[pos + end + 1..], nav_script)
                                    } else {
                                        format!("{}{}{}", base_tag, raw, nav_script)
                                    }
                                } else {
                                    format!("{}{}{}", base_tag, raw, nav_script)
                                }
                            }
                            _ => format!("<html><body><p style='color:red'>Failed to load: {}</p></body></html>", url),
                        };
                        eprintln!("[WebBrowser] link fetched {} bytes", html.len());
                        // Find which wb control to update (use first one for now)
                        let keys: Vec<String> = wb_html_poll.read().keys().cloned().collect();
                        if let Some(key) = keys.first() {
                            wb_html_poll.write().insert(key.clone(), html);
                        }
                    }
                }
            }
        });
    }

    // ── WebBrowser proxy handler ─────────────────────────────────────
    // Dioxus desktop blocks all non-dioxus:// navigations (including
    // iframe sub-frame loads) and opens http/https URLs in the system
    // browser.  We work around this by routing WebBrowser iframe URLs
    // through the dioxus:// custom protocol via this asset handler.
    // Requests to  /vybeweb/<encoded-url>  are fetched with curl and
    // the response HTML is returned inline.
    use_asset_handler("vybeweb", move |request, responder| {
        let path = request.uri().path();
        // Strip the leading /vybeweb/ prefix to get the target URL
        let target_url = path.strip_prefix("/vybeweb/").unwrap_or(path);
        let target_url = urlencoding::decode(target_url)
            .unwrap_or_else(|_| target_url.into())
            .to_string();

        if target_url.is_empty() || target_url == "about:blank" {
            responder.respond(
                Response::builder()
                    .header("Content-Type", "text/html")
                    .body(b"<html><body></body></html>".to_vec())
                    .unwrap(),
            );
            return;
        }

        // Fetch the URL using curl (available on macOS / Linux)
        let output = std::process::Command::new("curl")
            .args(&["-s", "-L", "-k", "--max-time", "10", &target_url])
            .output();

        match output {
            Ok(out) if out.status.success() => {
                // Inject a <base> tag so relative URLs resolve against the
                // original site, and add a target="_self" so links stay in
                // the iframe instead of opening in the system browser.
                let html = String::from_utf8_lossy(&out.stdout);
                let base_tag = format!(
                    "<base href=\"{}\" target=\"_self\">",
                    target_url
                );
                let patched = if let Some(pos) = html.to_lowercase().find("<head") {
                    // Insert right after the opening <head...>
                    if let Some(end) = html[pos..].find('>') {
                        format!("{}{}{}", &html[..pos + end + 1], base_tag, &html[pos + end + 1..])
                    } else {
                        format!("{}{}", base_tag, html)
                    }
                } else {
                    format!("{}{}", base_tag, html)
                };
                responder.respond(
                    Response::builder()
                        .header("Content-Type", "text/html; charset=utf-8")
                        .body(patched.into_bytes())
                        .unwrap(),
                );
            }
            _ => {
                let err_msg = format!(
                    "<html><body><p style='color:red'>Failed to load: {}</p></body></html>",
                    target_url
                );
                responder.respond(
                    Response::builder()
                        .header("Content-Type", "text/html")
                        .body(err_msg.into_bytes())
                        .unwrap(),
                );
            }
        }
    });

    // ── Initialize Runtime ──────────────────────────────────────────────
    use_effect(move || {
        if interpreter.read().is_none() && !*is_console_mode.read() {
            let project_read = rp.project.read();
            if let Some(proj) = project_read.as_ref() {
                // ── Sub Main mode (interactive console) ─────────────────
                if proj.starts_with_main() {
                    // Collect all code + resources as owned Strings so we can
                    // move them to the background thread.
                    let resource_entries = crate::runner::collect_resource_entries(proj);
                    let code_files: Vec<String> = proj.code_files.iter()
                        .map(|cf| cf.code.clone())
                        .collect();
                    let form_sources: Vec<(String, String)> = proj.forms.iter()
                        .map(|fm| {
                            let code = if fm.is_vbnet() {
                                format!("{}\n{}", fm.get_designer_code(), fm.get_user_code())
                            } else {
                                fm.get_user_code().to_string()
                            };
                            (fm.form.name.clone(), code)
                        })
                        .collect();
                    drop(project_read);

                    // Set up channels
                    let (msg_tx, msg_rx) = mpsc::channel::<ConsoleMessage>();
                    let (input_tx, input_rx) = mpsc::channel::<String>();

                    // Store channel endpoints for the UI
                    console_rx.set(Some(Rc::new(RefCell::new(Some(msg_rx)))));
                    console_input_tx.set(Some(Rc::new(RefCell::new(Some(input_tx)))));
                    is_console_mode.set(true);

                    // Spawn the interpreter on a background thread
                    std::thread::spawn(move || {
                        let mut interp = Interpreter::new();
                        interp.console_tx = Some(msg_tx.clone());
                        interp.console_input_rx = Some(input_rx);
                        interp.register_resource_entries(resource_entries);

                        // Load all code files
                        for code in &code_files {
                            if let Ok(program) = parse_program(code) {
                                let _ = interp.load_code_file(&program);
                            }
                        }
                        // Load form modules
                        for (name, code) in &form_sources {
                            if let Ok(program) = parse_program(code) {
                                let _ = interp.load_module(name, &program);
                            }
                        }

                        // Run Sub Main
                        match interp.call_procedure(&vybe_parser::ast::Identifier::new("main"), &[]) {
                            Ok(_) => { let _ = msg_tx.send(ConsoleMessage::Finished); }
                            Err(e) => { let _ = msg_tx.send(ConsoleMessage::Error(format!("{:?}", e))); }
                        }
                    });

                    return;
                }

                // ── Form mode ───────────────────────────────────────────
                if let Some(startup_form_module) = proj.get_startup_form() {
                    let form = startup_form_module.form.clone();
                    let form_code = if startup_form_module.is_vbnet() {
                        format!("{}\n{}", startup_form_module.get_designer_code(), startup_form_module.get_user_code())
                    } else {
                        startup_form_module.get_user_code().to_string()
                    };
                    drop(project_read);

                    runtime_form.set(Some(form.clone()));

                    let mut interp = Interpreter::new();

                    // Register resources from all project resource files + form resources
                    if let Some(proj) = rp.project.read().as_ref() {
                        let entries = crate::runner::collect_resource_entries(proj);
                        interp.register_resource_entries(entries);
                    }

                    // Load all code files (global scope + class definitions)
                    let project_read = rp.project.read();
                    if let Some(proj) = project_read.as_ref() {
                        for code_file in &proj.code_files {
                            if let Ok(program) = parse_program(&code_file.code) {
                                let _ = interp.run(&program);
                            }
                        }
                    }
                    drop(project_read);

                    // Now load the startup form code (registers the class)
                    match parse_program(&form_code) {
                        Ok(program) => {
                            parse_error.set(None);
                            if let Err(e) = interp.run(&program) {
                                parse_error.set(Some(format!("Runtime Load Error: {:?}", e)));
                            } else {
                                // ── Clean .NET-style form initialization ─────
                                // 1. Create the form instance via the interpreter.
                                //    This calls Sub New → InitializeComponent,
                                //    populating the instance fields with controls.
                                let form_class = form.name.clone();
                                let form_name_lower = form.name.to_lowercase();

                                match interp.create_class_instance(&form_class) {
                                    Ok(form_obj) => {
                                        // 2. Store the instance as __form_instance__
                                        //    in the global env so process_side_effects
                                        //    and event dispatch can find it.
                                        interp.env.define_global(
                                            "__form_instance__",
                                            Value::Object(form_obj.clone()),
                                        );

                                        // 3. Fill in any controls that the interpreter
                                        //    couldn't fully build from designer code.
                                        if let Some(frm) = runtime_form.read().as_ref() {
                                            ensure_controls_on_instance(frm, &form_obj);
                                        }

                                        // 4. Register each control as an env variable
                                        //    pointing to the SAME Rc, so `btn1.Text`
                                        //    (without `Me.`) also resolves correctly.
                                        {
                                            let obj_borrow = form_obj.borrow();
                                            for ctrl in &form.controls {
                                                let key = ctrl.name.to_lowercase();
                                                if let Some(ctrl_val) = obj_borrow.fields.get(&key) {
                                                    interp.env.define_global(&ctrl.name, ctrl_val.clone());
                                                }
                                            }
                                        }

                                        // 5. Pull instance state into the Form struct
                                        //    so the UI renders initial values correctly.
                                        if let Some(frm) = runtime_form.write().as_mut() {
                                            sync_instance_to_ui(frm, &form_obj);
                                        }

                                        // 6. Fire Form_Load / Me.Load handler
                                        let load_args = interp.make_event_handler_args(&form_class, "Load");
                                        let handlers = interp.get_event_handlers(&form_name_lower, "Me", "Load");
                                        if !handlers.is_empty() {
                                            for handler in handlers {
                                                let _ = interp.call_method_on_object(&form_obj, &handler, &load_args);
                                            }
                                        } else {
                                            // Fallback: try Form1_Load convention
                                            let conv = format!("{}_Load", form_class);
                                            let _ = interp.call_method_on_object(&form_obj, &conv, &load_args);
                                        }
                                    }
                                    Err(e) => {
                                        eprintln!("Form instance creation error: {:?}", e);
                                        // Fallback: still try Form_Load as a free sub
                                        let load_args = interp.make_event_handler_args("Form", "Load");
                                        let _ = interp.call_event_handler("Form_Load", &load_args);
                                    }
                                }

                                // 7. Process side-effects (MsgBox, form switches, etc.)
                                process_side_effects(&mut interp, rp, &mut runtime_form, &mut msgbox_content, &mut wb_html);

                                // 7b. Auto-fetch any WebBrowser controls that have a URL
                                //     set (from designer or Form_Load) but no content yet.
                                if let Some(frm) = runtime_form.read().as_ref() {
                                    let mut to_fetch: Vec<(String, String)> = Vec::new();
                                    for ctrl in &frm.controls {
                                        if ctrl.control_type == vybe_forms::ControlType::WebBrowser {
                                            let url = ctrl.properties.get_string("URL")
                                                .map(|s| s.to_string())
                                                .unwrap_or_default();
                                            let key = ctrl.name.to_lowercase();
                                            if !url.is_empty() && url != "about:blank"
                                                && !wb_html.read().contains_key(&key)
                                            {
                                                to_fetch.push((key, url));
                                            }
                                        }
                                    }
                                    for (key, url) in to_fetch {
                                        eprintln!("[WebBrowser] auto-fetching initial URL: {}", url);
                                        let output = std::process::Command::new("curl")
                                            .args(&["-s", "-L", "-k", "--max-time", "10", &url])
                                            .output();
                                        let html = match output {
                                            Ok(out) if out.status.success() => {
                                                let raw = String::from_utf8_lossy(&out.stdout).to_string();
                                                let base_tag = format!("<base href=\"{}\" target=\"_self\">", url);
                                                let nav_script = "<script>\
                                                    document.addEventListener('click',function(e){\
                                                      var a=e.target.closest('a');\
                                                      if(a&&a.href){\
                                                        e.preventDefault();\
                                                        e.stopPropagation();\
                                                        window.__vybe_nav=a.href;\
                                                      }\
                                                    },true);\
                                                    </script>";
                                                if let Some(pos) = raw.to_lowercase().find("<head") {
                                                    if let Some(end) = raw[pos..].find('>') {
                                                        format!("{}{}{}{}", &raw[..pos + end + 1], base_tag, &raw[pos + end + 1..], nav_script)
                                                    } else {
                                                        format!("{}{}{}", base_tag, raw, nav_script)
                                                    }
                                                } else {
                                                    format!("{}{}{}", base_tag, raw, nav_script)
                                                }
                                            }
                                            _ => format!("<html><body><p style='color:red'>Failed to load: {}</p></body></html>", url),
                                        };
                                        eprintln!("[WebBrowser] auto-fetched {} bytes", html.len());
                                        wb_html.write().insert(key, html);
                                    }
                                }

                                // 8. Fire Shown event
                                if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                    let shown_form_name = runtime_form.read().as_ref().map(|f| f.name.clone());
                                    if let Some(fname) = shown_form_name {
                                        let shown_args = interp.make_event_handler_args(&fname, "Shown");
                                        let fname_lower = fname.to_lowercase();
                                        let handlers = interp.get_event_handlers(&fname_lower, "Me", "Shown");
                                        if !handlers.is_empty() {
                                            for handler in handlers {
                                                let _ = interp.call_method_on_object(&form_obj, &handler, &shown_args);
                                            }
                                        } else {
                                            let _ = interp.call_method_on_object(&form_obj, &format!("{}_Shown", fname), &shown_args);
                                        }
                                        process_side_effects(&mut interp, rp, &mut runtime_form, &mut msgbox_content, &mut wb_html);
                                    }
                                }

                                // 9. Post-init binding refresh: re-emit current position for
                                //    every BindingSource so nav + bound controls always show
                                //    the correct initial state regardless of init ordering.
                                let bs_objects: Vec<std::rc::Rc<std::cell::RefCell<vybe_runtime::value::ObjectData>>> = {
                                    if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                                        let form_borrow = form_obj.borrow();
                                        form_borrow.fields.values()
                                            .filter_map(|val| {
                                                if let Value::Object(ctrl_obj) = val {
                                                    let is_bs = ctrl_obj.borrow().fields
                                                        .get("__type")
                                                        .map(|v| v.as_string() == "BindingSource")
                                                        .unwrap_or(false);
                                                    if is_bs { Some(ctrl_obj.clone()) } else { None }
                                                } else { None }
                                            })
                                            .collect()
                                    } else { Vec::new() }
                                };
                                for ctrl_obj in bs_objects {
                                    let (bs_name, position) = {
                                        let borrow = ctrl_obj.borrow();
                                        let name = borrow.fields.get("name").map(|v| v.as_string()).unwrap_or_default();
                                        let pos = borrow.fields.get("position")
                                            .and_then(|v| if let Value::Integer(i) = v { Some(*i) } else { None })
                                            .unwrap_or(0);
                                        (name, pos)
                                    };
                                    if bs_name.is_empty() { continue; }
                                    let count = interp.binding_source_row_count_filtered(&Value::Object(ctrl_obj.clone()));
                                    if count > 0 {
                                        let ds = ctrl_obj.borrow().fields.get("__datasource").cloned().unwrap_or(Value::Nothing);
                                        interp.side_effects.push_back(RuntimeSideEffect::BindingPositionChanged {
                                            binding_source_name: bs_name,
                                            position,
                                            count,
                                        });
                                        // Refresh DataBindings.Add bound controls (TextBox, Label, CheckBox, etc.)
                                        interp.refresh_bindings_filtered(&ctrl_obj, &ds, position);
                                        // Emit DataSourceChanged for controls bound via DataSource = bs
                                        // (DataGridView, ComboBox, ListBox, etc.)
                                        let bound_controls: Vec<String> = ctrl_obj.borrow()
                                            .fields.get("__bound_controls")
                                            .and_then(|v| if let Value::Array(arr) = v {
                                                Some(arr.iter().filter_map(|v| if let Value::String(s) = v { Some(s.clone()) } else { None }).collect())
                                            } else { None })
                                            .unwrap_or_default();
                                        let bs_val = Value::Object(ctrl_obj.clone());
                                        for bound_ctrl_name in bound_controls {
                                            let (columns, rows) = interp.get_datasource_table_data_filtered(&bs_val);
                                            interp.side_effects.push_back(RuntimeSideEffect::DataSourceChanged {
                                                control_name: bound_ctrl_name,
                                                columns,
                                                rows,
                                            });
                                        }
                                    }
                                }
                                if !interp.side_effects.is_empty() {
                                    process_side_effects(&mut interp, rp, &mut runtime_form, &mut msgbox_content, &mut wb_html);
                                }
                            }
                        }
                        Err(e) => {
                            parse_error.set(Some(format!("Parse Error: {:?}", e)));
                        }
                    }
                    interpreter.set(Some(interp));
                } else {
                    println!("No startup form defined in project");
                }
            }
        }
    });

    // ── Event handler ───────────────────────────────────────────────────
    let mut handle_event = move |control_name: String, event_name: String, event_data: Option<vybe_runtime::EventData>| {
        if parse_error.read().is_some() {
            return;
        }
        if msgbox_content.read().is_some() {
            return;
        }
        // Re-entrancy guard: if we're already handling an event (e.g. a modal
        // dialog is pumping the macOS event loop), reject new events to avoid
        // deadlocking on the interpreter write lock.
        if *handling_event.read() {
            return;
        }

        // Passive mouse-tracking events: dispatch to handler if one exists
        // but skip the full UI sync / side-effects to avoid infinite
        // mousemove → re-render → mousemove loops.
        let is_passive_mouse = matches!(
            event_name.as_str(),
            "MouseMove" | "MouseEnter" | "MouseLeave"
        );

        if is_passive_mouse {
            if *handling_event.read() {
                return;
            }
            if let Some(interp) = interpreter.write().as_mut() {
                if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                    let form_class = runtime_form.peek().as_ref()
                        .map(|f| f.name.to_lowercase()).unwrap_or_default();
                    let event_args = interp.make_event_handler_args_with_data(&control_name, &event_name, event_data.as_ref());
                    let handlers = interp.get_event_handlers(&form_class, &control_name, &event_name);
                    for method_name in handlers {
                        let _ = interp.call_method_on_object(&form_obj, &method_name, &event_args);
                    }
                }
            }
            return;
        }

        handling_event.set(true);
        if let Some(interp) = interpreter.write().as_mut() {
            // ── Pre-event sync: push UI state → instance ──────────
            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                if let Some(frm) = runtime_form.read().as_ref() {
                    sync_ui_to_instance(frm, &form_obj);
                }
            }

            // ── Dispatch event ─────────────────────────────────────
            let mut executed = false;

            if let Ok(Value::Object(form_obj)) = interp.env.get("__form_instance__") {
                let form_class = runtime_form.read().as_ref()
                    .map(|f| f.name.to_lowercase()).unwrap_or_default();
                let event_args = interp.make_event_handler_args_with_data(
                    &control_name, &event_name, event_data.as_ref(),
                );

                // 1. Try Handles clause AND AddHandler (e.g. `Handles btn1.Click` or `AddHandler ...`)
                let handlers = interp.get_event_handlers(&form_class, &control_name, &event_name);
                for method_name in handlers {
                    if interp.call_method_on_object(&form_obj, &method_name, &event_args).is_ok() {
                        executed = true;
                    }
                }

                // 2. Try conventional name (e.g. `btn1_Click`)
                if !executed {
                    let conv_name = format!("{}_{}", control_name, event_name);
                    if interp.call_method_on_object(&form_obj, &conv_name, &event_args).is_ok() {
                        executed = true;
                    }
                }

                // 3. For form-level events, also try Form1_Load style
                if !executed {
                    if let Some(frm) = runtime_form.read().as_ref() {
                        if control_name.eq_ignore_ascii_case(&frm.name) {
                            let alt = format!("{}_{}", frm.name, event_name);
                            if interp.call_method_on_object(&form_obj, &alt, &event_args).is_ok() {
                                executed = true;
                            }
                        }
                    }
                }
            }

            // 4. Fallback: try as a free sub (non-class handler)
            if !executed {
                let handler_lower = format!("{}_{}", control_name, event_name).to_lowercase();
                let handler_key = interp
                    .subs
                    .keys()
                    .find(|key| *key == &handler_lower || key.ends_with(&format!(".{}", handler_lower)))
                    .cloned();
                if let Some(key) = handler_key {
                    let event_args = interp.make_event_handler_args_with_data(&control_name, &event_name, event_data.as_ref());
                    let _ = interp.call_event_handler(&key, &event_args);
                }
            }

            // ── BindingNavigator auto-delegation ──────────────────
            if let Some(frm) = runtime_form.read().as_ref() {
                let is_nav_event = matches!(
                    event_name.as_str(),
                    "MoveFirst" | "MoveNext" | "MovePrevious" | "MoveLast" | "AddNew" | "Delete"
                );
                if is_nav_event {
                    if let Some(ctrl) = frm.controls.iter().find(|c| c.name == control_name) {
                        if matches!(ctrl.control_type, vybe_forms::ControlType::BindingNavigator) {
                            let bs_name = ctrl.properties.get_string("BindingSource").unwrap_or_default();
                            if !bs_name.is_empty() {
                                if let Ok(Value::Object(_form_obj)) = interp.env.get("__form_instance__") {
                                    // Navigate via the form instance
                                    let nav_script = format!("__form_instance__.{}.{}()", bs_name, event_name);
                                    if let Ok(prog) = parse_program(&nav_script) {
                                        let _ = interp.load_module("NavAction", &prog);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            process_side_effects(interp, rp, &mut runtime_form, &mut msgbox_content, &mut wb_html);
        }
        handling_event.set(false);
    };

    // ── Console message polling ─────────────────────────────────────────
    // In console mode, poll the channel from the interpreter thread every
    // 50 ms and update signals accordingly.
    use_future(move || async move {
        loop {
            tokio::time::sleep(std::time::Duration::from_millis(50)).await;
            if !*is_console_mode.read() { continue; }

            // Drain all available messages
            let rx_opt = console_rx.read().clone();
            if let Some(rx_rc) = rx_opt {
                let rx_ref = rx_rc.borrow();
                if let Some(rx) = rx_ref.as_ref() {
                    loop {
                        match rx.try_recv() {
                            Ok(ConsoleMessage::Output { text, fg, bg }) => {
                                let escaped = html_escape(&text);
                                // Convert newlines to <br> since we use innerHTML
                                let with_br = escaped.replace('\n', "<br>");
                                let fg_css = console_color_to_css(fg);
                                if bg == 0 {
                                    // No background (transparent) for default
                                    console_output.write().push_str(
                                        &format!("<span style=\"color:{}\">{}</span>", fg_css, with_br)
                                    );
                                } else {
                                    let bg_css = console_color_to_css(bg);
                                    console_output.write().push_str(
                                        &format!("<span style=\"color:{};background:{}\">{}</span>", fg_css, bg_css, with_br)
                                    );
                                }
                            }
                            Ok(ConsoleMessage::Clear) => {
                                console_output.set(String::new());
                            }
                            Ok(ConsoleMessage::InputRequest) => {
                                console_waiting_input.set(true);
                            }
                            Ok(ConsoleMessage::Finished) => {
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                            Ok(ConsoleMessage::Error(e)) => {
                                console_output.write().push_str(
                                    &format!("<br><span style=\"color:#e74856\">--- Error: {} ---</span><br>", html_escape(&e))
                                );
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                            Err(mpsc::TryRecvError::Empty) => break,
                            Err(mpsc::TryRecvError::Disconnected) => {
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                        }
                    }
                }
            }
        }
    });

    // ── Render ───────────────────────────────────────────────────────────
    let form_opt = runtime_form.read().clone();
    let msgbox_visible = msgbox_content.read().is_some();
    let msgbox_text = msgbox_content.read().clone().unwrap_or_default();
    let error_text = parse_error.read().clone();

    rsx! {
        div {
            class: "runtime-panel",
            style: "flex: 1; background: #e0e0e0; display: flex; align-items: center; justify-content: center; overflow: auto; position: relative;",

            {
                if *is_console_mode.read() {
                    // ── Interactive Console UI ──────────────────────────
                    let output_text = console_output.read().clone();
                    let waiting = *console_waiting_input.read();
                    let finished = *console_finished.read();
                    let input_val = console_input_line.read().clone();

                    rsx! {
                        div {
                            style: "
                                width: 700px;
                                height: 480px;
                                background: #1e1e1e;
                                color: #d4d4d4;
                                border: 1px solid #444;
                                box-shadow: 0 0 10px rgba(0,0,0,0.5);
                                display: flex;
                                flex-direction: column;
                                font-family: 'Consolas', 'Courier New', monospace;
                                font-size: 14px;
                            ",

                            // Title bar
                            div {
                                style: "
                                    background: #2d2d2d;
                                    color: #ccc;
                                    padding: 4px 10px;
                                    font-size: 12px;
                                    border-bottom: 1px solid #444;
                                    user-select: none;
                                ",
                                "Console Output"
                            }

                            // Output area
                            div {
                                id: "console-output",
                                style: "
                                    flex: 1;
                                    overflow-y: auto;
                                    padding: 8px 10px;
                                    white-space: pre-wrap;
                                    word-break: break-all;
                                ",
                                dangerous_inner_html: "{output_text}",
                            }

                            // Input area (shown when waiting for input, or always visible with prompt)
                            if waiting && !finished {
                                div {
                                    style: "
                                        display: flex;
                                        border-top: 1px solid #444;
                                        background: #252526;
                                    ",
                                    span {
                                        style: "padding: 6px 4px 6px 10px; color: #569cd6;",
                                        ">"
                                    }
                                    input {
                                        id: "console-input",
                                        style: "
                                            flex: 1;
                                            background: #252526;
                                            color: #d4d4d4;
                                            border: none;
                                            outline: none;
                                            padding: 6px 8px;
                                            font-family: 'Consolas', 'Courier New', monospace;
                                            font-size: 14px;
                                        ",
                                        value: "{input_val}",
                                        autofocus: true,
                                        oninput: move |evt| {
                                            console_input_line.set(evt.value().clone());
                                        },
                                        onkeypress: move |evt: KeyboardEvent| {
                                            if evt.key() == Key::Enter {
                                                let line = console_input_line.read().clone();
                                                // Echo the input to the output
                                                console_output.write().push_str(&format!("{}\n", line));
                                                // Send to the interpreter thread
                                                if let Some(tx_rc) = console_input_tx.read().as_ref() {
                                                    if let Some(tx) = tx_rc.borrow().as_ref() {
                                                        let _ = tx.send(line);
                                                    }
                                                }
                                                console_input_line.set(String::new());
                                                console_waiting_input.set(false);
                                            }
                                        },
                                    }
                                }
                            }

                            // Status bar
                            div {
                                style: "
                                    background: #007acc;
                                    color: white;
                                    padding: 2px 10px;
                                    font-size: 11px;
                                ",
                                if finished {
                                    "Program finished"
                                } else if waiting {
                                    "Waiting for input..."
                                } else {
                                    "Running..."
                                }
                            }
                        }
                    }
                } else if let Some(err) = error_text {
                    rsx! {
                        div {
                            style: "color: red; padding: 20px; background: #fee; border: 1px solid red;",
                            "Error: {err}"
                        }
                    }
                } else if let Some(form) = form_opt {
                    let width = form.width;
                    let height = form.height;
                    let caption = form.text.clone();
                    let form_back = form.back_color.clone().unwrap_or_else(|| "#f8fafc".to_string());
                    let form_fore = form.fore_color.clone().unwrap_or_else(|| "#0f172a".to_string());
                    let form_font = form.font.clone().unwrap_or_else(|| "Segoe UI, 12px".to_string());

                    rsx! {
                        div {
                            style: "
                                position: relative;
                                width: {width}px;
                                height: {height}px;
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
                                span { style: "flex: 1;", "{caption}" }
                                // Close button
                                button {
                                    style: "background: transparent; border: none; color: white; font-size: 16px; cursor: pointer; padding: 0 4px; margin-left: auto; font-weight: bold; line-height: 1;",
                                    onclick: move |_| {
                                        let form_name_close = form.name.clone();
                                        handle_event(form_name_close, "FormClosing".to_string(), None);
                                    },
                                    "\u{00D7}"
                                }
                            }

                            // Controls (skip non-visual components)
                            ControlTree {
                                form: form.clone(),
                                parent_id: None,
                                interpreter: interpreter,
                                wb_html: wb_html,
                                runtime_form: runtime_form,
                                on_handle_event: move |(name, evt_name, data)| handle_event(name, evt_name, data),
                            }
                            // MsgBox overlay
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
                                            div {
                                                style: "background: #0078d4; color: white; padding: 4px 8px; font-weight: bold;",
                                                "Project1"
                                            }
                                            div {
                                                style: "padding: 20px; text-align: center; color: black;",
                                                "{msgbox_text}"
                                            }
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
