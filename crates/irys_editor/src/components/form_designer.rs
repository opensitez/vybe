use dioxus::prelude::*;
use crate::app_state::AppState;
use irys_forms::{Control, ControlType, Bounds};
use uuid::Uuid;
use std::collections::HashMap;

#[derive(Clone, Copy, Debug, PartialEq)]
enum HandlePosition {
    TopLeft, Top, TopRight,
    Left, Right,
    BottomLeft, Bottom, BottomRight
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum DragMode {
    Pending, // New state: waiting for move threshold
    Move,
    Resize(HandlePosition)
}

#[derive(Clone, Debug)]
struct DragState {
    start_x: i32,
    start_y: i32,
    // Current mouse pos to track delta for pending state
    // Actually we can just compare against start_x/y
    initial_bounds: Bounds,
    mode: DragMode,
}

#[component]
pub fn FormDesigner() -> Element {
    let state = use_context::<AppState>();
    let form_opt = state.get_current_form();
    let mut selected_tool = state.selected_tool;
    let mut selected_control = state.selected_control;
    
    // Local state for dragging
    let mut drag_state = use_signal(|| None::<DragState>);
    let mut drop_target = use_signal(|| None::<Option<Uuid>>); // None = Form, Some(id) = Container, None in Option means nothing

    // Helper to organize controls by parent
    // We compute this on every render which is fine for small forms
    // Map: Option<ParentId> -> Vec<&Control>
    let mut hierarchy: HashMap<Option<Uuid>, Vec<Control>> = HashMap::new();
    let mut control_map: HashMap<Uuid, Control> = HashMap::new();
    
    if let Some(form) = &form_opt {
        for control in &form.controls {
            hierarchy.entry(control.parent_id)
                .or_default()
                .push(control.clone());
            control_map.insert(control.id, control.clone());
        }
    }


    // Move handler (Global for the designer area)
    let handle_move = move |evt: MouseEvent| {
        let current_ds = drag_state.read().clone();

        if let Some(mut ds) = current_ds { 
            if let Some(_) = state.get_current_form() {
                let control_id_opt = *selected_control.read();

                if let Some(control_id) = control_id_opt {
                    let client_x = evt.client_coordinates().x as i32;
                    let client_y = evt.client_coordinates().y as i32;
                    // Calculate delta
                    let delta_x = client_x - ds.start_x;
                    let delta_y = client_y - ds.start_y;

                    // Handle Pending State
                    if matches!(ds.mode, DragMode::Pending) {
                        if delta_x.abs() > 5 || delta_y.abs() > 5 {
                            // Threshold passed, switch to Move
                            ds.mode = DragMode::Move;
                            drag_state.set(Some(ds)); // Update state to trigger pointer-events: none
                        }
                        return; // Don't move yet
                    }
                    
                    let mut new_bounds = ds.initial_bounds;

                    match ds.mode {
                        DragMode::Move => {
                            new_bounds.x += delta_x;
                            new_bounds.y += delta_y;
                            
                            // Snap to grid (10px)
                            new_bounds.x = (new_bounds.x / 10) * 10;
                            new_bounds.y = (new_bounds.y / 10) * 10;
                        },
                        DragMode::Resize(handle) => {
                            // Logic for resizing based on handle
                            // We construct new bounds based on delta
                            // Snap delta to grid?
                            
                            // Simple approach: apply delta to edges
                            match handle {
                                HandlePosition::Right => new_bounds.width = (new_bounds.width + delta_x).max(10),
                                HandlePosition::Bottom => new_bounds.height = (new_bounds.height + delta_y).max(10),
                                HandlePosition::BottomRight => {
                                    new_bounds.width = (new_bounds.width + delta_x).max(10);
                                    new_bounds.height = (new_bounds.height + delta_y).max(10);
                                },
                                HandlePosition::Left => {
                                    let old_right = new_bounds.x + new_bounds.width;
                                    new_bounds.x = (new_bounds.x + delta_x).min(old_right - 10);
                                    new_bounds.width = old_right - new_bounds.x;
                                },
                                HandlePosition::Top => {
                                    let old_bottom = new_bounds.y + new_bounds.height;
                                    new_bounds.y = (new_bounds.y + delta_y).min(old_bottom - 10);
                                    new_bounds.height = old_bottom - new_bounds.y;
                                },
                                // ... implement others as needed, simplified for main ones
                                _ => {} 
                            }
                            
                            // Align to grid
                            new_bounds.width = (new_bounds.width / 10).max(1) * 10;
                            new_bounds.height = (new_bounds.height / 10).max(1) * 10;
                            if matches!(handle, HandlePosition::Left) { new_bounds.x = (new_bounds.x / 10) * 10; }
                            if matches!(handle, HandlePosition::Top) { new_bounds.y = (new_bounds.y / 10) * 10; }
                        },
                        _ => {}
                    }
                    
                    // Update control in state
                    state.update_control_geometry(
                        control_id, 
                        new_bounds.x, 
                        new_bounds.y, 
                        new_bounds.width, 
                        new_bounds.height
                    );
                }
            }
        }
    };

    let handle_up = move |_| {
        let mut should_reparent = None;
        
        if let Some(ds) = drag_state.read().as_ref() {
            if matches!(ds.mode, DragMode::Move) {
                if let Some(control_id) = *selected_control.read() {
                     if let Some(target_opt) = *drop_target.read() {
                        should_reparent = Some((control_id, target_opt));
                     }
                }
            }
        }
        
        if let Some((control_id, target_opt)) = should_reparent {
             state.reparent_control(control_id, target_opt);
        }
        
        if drag_state.read().is_some() {
            drag_state.set(None);
        }
    };

    let handle_keydown = move |evt: KeyboardEvent| {
        let key = evt.key();
        let modifiers = evt.modifiers();
        let is_ctrl_or_meta = modifiers.contains(Modifiers::CONTROL) || modifiers.contains(Modifiers::META);

        match key {
            Key::Delete | Key::Backspace => {
                state.delete_selected_control();
            }
            Key::Character(ref c) if is_ctrl_or_meta && (c == "c" || c == "C") => {
                state.copy_selected_control();
            }
            Key::Character(ref c) if is_ctrl_or_meta && (c == "x" || c == "X") => {
                state.cut_selected_control();
            }
            Key::Character(ref c) if is_ctrl_or_meta && (c == "v" || c == "V") => {
                state.paste_control();
            }
            _ => {}
        }
    };

    rsx! {
        div {
            class: "form-designer",
            style: "flex: 1; background: #e0e0e0; position: relative; overflow: auto; outline: none;",
            tabindex: "0",
            onkeydown: handle_keydown,
            onmousemove: handle_move,
            onmouseup: handle_up,
            
            if let Some(form) = form_opt {
                {
                    let form_width = form.width;
                    let form_height = form.height;
                    let form_caption = form.caption.clone();
                    let form_back = form.back_color.clone().unwrap_or_else(|| "#f8fafc".to_string());
                    let form_fore = form.fore_color.clone().unwrap_or_else(|| "#0f172a".to_string());
                    let form_font = form.font.clone().unwrap_or_else(|| "Segoe UI, 12px".to_string());
                    
                    rsx! {
                        // Form canvas
                        div {
                            style: "
                                position: relative;
                                width: {form_width}px;
                                height: {form_height}px;
                                background: {form_back};
                                margin: 20px;
                                border: 1px solid #999;
                                box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
                                background-image: 
                                    linear-gradient(0deg, transparent 24%, rgba(0,0,0,.05) 25%, rgba(0,0,0,.05) 26%, transparent 27%, transparent 74%, rgba(0,0,0,.05) 75%, rgba(0,0,0,.05) 76%, transparent 77%, transparent),
                                    linear-gradient(90deg, transparent 24%, rgba(0,0,0,.05) 25%, rgba(0,0,0,.05) 26%, transparent 27%, transparent 74%, rgba(0,0,0,.05) 75%, rgba(0,0,0,.05) 76%, transparent 77%, transparent);
                                background-size: 20px 20px;
                                color: {form_fore};
                                font: {form_font};
                            ",
                            onmouseover: move |_| { drop_target.set(Some(None)); }, // Target is Form (None parent)
                            onclick: move |evt| {
                                let tool_opt = *selected_tool.read();
                                if let Some(tool) = tool_opt {
                                     // Non-visual components don't need coordinates
                                     if tool.is_non_visual() {
                                         state.add_control_at(tool, 0, 0);
                                         selected_tool.set(None);
                                         return;
                                     }
                                     let x = evt.data.element_coordinates().x as i32;
                                     let y = evt.data.element_coordinates().y as i32;
                                     let grid_x = (x / 10) * 10;
                                     let grid_y = (y / 10) * 10;

                                     state.add_control_at(tool, grid_x, grid_y);
                                     selected_tool.set(None);
                                } else {
                                     // Only deselect if not dragging (handled by stop_propagation on controls)
                                     selected_control.set(None);
                                }
                            },
                            
                             // Form title bar
                            div {
                                style: "
                                    position: absolute;
                                    top: 0; left: 0; right: 0; height: 30px;
                                    background: linear-gradient(to bottom, #0078d4, #005a9e);
                                    color: {form_fore};
                                    padding: 4px 8px;
                                    font-weight: bold;
                                    display: flex; align-items: center;
                                ",
                                "{form_caption}"
                            }
                            
                            
                            // Root controls
                            RecursiveControls { 
                                parent_id: None, 
                                hierarchy: hierarchy.clone(),
                                selected_control: selected_control,
                                drag_state: drag_state,
                                drop_target: drop_target,
                                parent_is_dragging: false,
                                depth: 0
                            }
                        }

                        // Component tray (non-visual data components)
                        {
                            let non_visual: Vec<Control> = form.controls.iter()
                                .filter(|c| c.control_type.is_non_visual())
                                .cloned()
                                .collect();

                            if !non_visual.is_empty() {
                                rsx! {
                                    div {
                                        style: "
                                            margin: 0 20px 20px 20px;
                                            width: {form_width}px;
                                            background: #f0f0f0;
                                            border: 1px solid #999;
                                            border-top: 2px solid #bbb;
                                            padding: 6px 8px;
                                            display: flex;
                                            flex-wrap: wrap;
                                            gap: 8px;
                                            min-height: 48px;
                                        ",

                                        // Tray label
                                        div {
                                            style: "width: 100%; font-size: 10px; color: #777; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px;",
                                            "Component Tray"
                                        }

                                        for comp in non_visual {
                                            {
                                                let comp_id = comp.id;
                                                let comp_name = comp.name.clone();
                                                let is_sel = *selected_control.read() == Some(comp_id);
                                                let sel_bg = if is_sel { "#cce4f7" } else { "#e8e8e8" };
                                                let sel_border = if is_sel { "2px solid #0078d4" } else { "1px solid #bbb" };
                                                let icon = match comp.control_type {
                                                    ControlType::BindingSourceComponent => "🔗",
                                                    ControlType::DataSetComponent => "🗄️",
                                                    ControlType::DataTableComponent => "📋",
                                                    ControlType::DataAdapterComponent => "🔌",
                                                    _ => "📦",
                                                };
                                                rsx! {
                                                    div {
                                                        key: "{comp_id}",
                                                        style: "
                                                            display: flex; flex-direction: column; align-items: center;
                                                            padding: 4px 10px; background: {sel_bg}; border: {sel_border};
                                                            border-radius: 4px; cursor: pointer; min-width: 64px;
                                                        ",
                                                        onclick: move |evt| {
                                                            evt.stop_propagation();
                                                            selected_control.set(Some(comp_id));
                                                        },
                                                        div { style: "font-size: 20px;", "{icon}" }
                                                        div { style: "font-size: 10px; color: #333; white-space: nowrap;", "{comp_name}" }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                rsx! {}
                            }
                        }
                    }
                }
            } else {
                div { style: "padding: 20px; color: #999;", "No form selected" }
            }
        }
    }
}

#[component]
fn RecursiveControls(
    parent_id: Option<Uuid>, 
    hierarchy: HashMap<Option<Uuid>, Vec<Control>>,
    selected_control: Signal<Option<Uuid>>,
    drag_state: Signal<Option<DragState>>,
    drop_target: Signal<Option<Option<Uuid>>>,
    parent_is_dragging: bool,
    depth: usize
) -> Element {
    if depth > 50 {
        return rsx! { div { "Max depth reached" } };
    }

    // Get children for this parent, filtering out non-visual components
    let children: Vec<Control> = hierarchy.get(&parent_id).cloned().unwrap_or_default()
        .into_iter()
        .filter(|c| !c.control_type.is_non_visual())
        .collect();
    
    rsx! {
        for control in children {
            RecursiveControlItem {
                key: "{control.id}",
                control: control,
                hierarchy: hierarchy.clone(),
                selected_control: selected_control,
                drag_state: drag_state,
                drop_target: drop_target,
                parent_is_dragging: parent_is_dragging,
                depth: depth
            }
        }
    }
}


#[component]
fn RecursiveControlItem(
    control: Control, 
    hierarchy: HashMap<Option<Uuid>, Vec<Control>>,
    selected_control: Signal<Option<Uuid>>,
    drag_state: Signal<Option<DragState>>,
    drop_target: Signal<Option<Option<Uuid>>>,
    parent_is_dragging: bool,
    depth: usize
) -> Element {
    let control_id = control.id;
    let display_name = control.display_name();
    let is_selected = *selected_control.read() == Some(control_id);
    let border_style = if is_selected { "2px dashed #0078d4" } else { "none" };

    // Check if dragging (Move mode only for pointer-events)
    let is_self_dragging = drag_state.read().as_ref().map_or(false, |ds| matches!(ds.mode, DragMode::Move) && is_selected);
    
    // Effectively dragging if self is dragging OR parent is dragging
    let is_essentially_dragging = is_self_dragging || parent_is_dragging;
    
    let pointer_events = if is_essentially_dragging { "none" } else { "auto" };

    let handle_down = move |evt: MouseEvent| {
        evt.stop_propagation(); 
        if !is_selected {
            selected_control.set(Some(control_id));
        }
        drag_state.set(Some(DragState {
            start_x: evt.client_coordinates().x as i32,
            start_y: evt.client_coordinates().y as i32,
            initial_bounds: control.bounds,
            mode: DragMode::Pending // Start in Pending
        }));
    };

    rsx! {
        div {
            key: "{control_id}",
            title: "{display_name}",
            style: "
                position: absolute;
                left: {control.bounds.x}px;
                top: {control.bounds.y}px;
                width: {control.bounds.width}px;
                height: {control.bounds.height}px;
                border: {border_style};
                cursor: default;
                user-select: none;
                pointer-events: {pointer_events};
            ",
            onmousedown: handle_down,
            onclick: move |evt| { evt.stop_propagation(); }, // Prevent deselecting when clicked
            
            ControlContent {
                control: control.clone(),
                hierarchy: hierarchy.clone(),
                selected_control: selected_control,
                drag_state: drag_state,
                drop_target: drop_target,
                is_dragging: is_essentially_dragging,
                depth: depth
            }

            if is_selected {
                ResizeHandles { 
                    control_bounds: control.bounds, 
                    drag_state: drag_state
                }
            }
        }
    }
}

#[component]
fn ControlContent(
    control: Control,
    hierarchy: HashMap<Option<Uuid>, Vec<Control>>,
    selected_control: Signal<Option<Uuid>>,
    drag_state: Signal<Option<DragState>>,
    drop_target: Signal<Option<Option<Uuid>>>,
    is_dragging: bool,
    depth: usize
) -> Element {
    let _state = use_context::<AppState>();
    let control_id = control.id;
    let control_type = control.control_type;
    let caption = control.get_caption().map(|s| s.to_string()).unwrap_or(control.name.clone());

    match control_type {
        ControlType::RichTextBox => rsx! {
            {
                let html = control.properties.get_string("HTML")
                    .map(|s| s.to_string())
                    .or_else(|| control.get_text().map(|s| s.to_string()))
                    .unwrap_or_default();
                let back = control.get_back_color().map(|s| s.to_string()).unwrap_or_else(|| "#f8fafc".to_string());
                let fore = control.get_fore_color().map(|s| s.to_string()).unwrap_or_else(|| "#0f172a".to_string());
                let font = control.get_font().map(|s| s.to_string()).unwrap_or_else(|| "Segoe UI, 12px".to_string());
                rsx! {
                    div {
                        style: "width: 100%; height: 100%; padding: 8px; overflow: auto; border: 1px inset #999; background: {back}; color: {fore}; font: {font}; pointer-events: none;",
                        dangerous_inner_html: "{html}",
                    }
                }
            }
        },
        ControlType::Frame | ControlType::PictureBox | ControlType::WebBrowser | ControlType::Panel => {
            let back = control.get_back_color().map(|s| s.to_string()).unwrap_or_else(|| "#f8fafc".to_string());
            let fore = control.get_fore_color().map(|s| s.to_string()).unwrap_or_else(|| "#0f172a".to_string());
            let font = control.get_font().map(|s| s.to_string()).unwrap_or_else(|| "Segoe UI, 12px".to_string());
            rsx! {
                div {
                    style: "width: 100%; height: 100%; border: 1px solid #999; position: relative; background: {back}; color: {fore}; font: {font};",
                    onmouseover: move |evt| {
                        evt.stop_propagation();
                        drop_target.set(Some(Some(control_id)));
                    },
                    if control_type == ControlType::Frame {
                        div {
                            style: "position: absolute; top: -8px; left: 8px; background: {back}; padding: 0 4px; font-size: 11px; color: {fore};",
                            "{caption}"
                        }
                    }
                    RecursiveControls {
                        parent_id: Some(control_id),
                        hierarchy: hierarchy,
                        selected_control: selected_control,
                        drag_state: drag_state,
                        drop_target: drop_target,
                        parent_is_dragging: is_dragging,
                        depth: depth + 1
                    }
                }
            }
        },
        _ => rsx! {
            div {
                style: "width: 100%; height: 100%; pointer-events: none;", 
                ControlVisuals { control: control.clone() }
            }
        }
    }
}

#[component]
fn ControlVisuals(control: Control) -> Element {
    let caption = control.get_caption().map(|s| s.to_string()).unwrap_or(control.name.clone());
    let text = control.get_text().map(|s| s.to_string()).unwrap_or_default();
    let back = control.get_back_color().map(|s| s.to_string()).unwrap_or_else(|| "#f8fafc".to_string());
    let fore = control.get_fore_color().map(|s| s.to_string()).unwrap_or_else(|| "#0f172a".to_string());
    let font = control.get_font().map(|s| s.to_string()).unwrap_or_else(|| "Segoe UI, 12px".to_string());
    
    match control.control_type {
        ControlType::Button => rsx! {
            div {
                style: "width: 100%; height: 100%; border: 1px outset #999; display: flex; align-items: center; justify-content: center; padding: 4px 6px; font: {font}; color: {fore}; background: {back};",
                "{caption}"
            }
        },
        ControlType::Label => rsx! {
            div {
                 style: "width: 100%; height: 100%; padding: 2px; font: {font}; color: {fore}; overflow: hidden; background: {back};",
                "{caption}"
            }
        },
        ControlType::TextBox => rsx! {
             div {
                style: "width: 100%; height: 100%; border: 1px inset #999; padding: 4px; font: {font}; color: {fore}; background: {back}; overflow: hidden;",
                "{text}"
             }
        },
        ControlType::CheckBox => rsx! {
             div {
                style: "display: flex; align-items: center; gap: 6px; font: {font}; color: {fore}; background: {back}; padding: 2px 4px;",
                div { style: "width: 13px; height: 13px; border: 1px solid #999; background: white;" }
                "{caption}"
             }
        },
        ControlType::RadioButton => rsx! {
             div {
                style: "display: flex; align-items: center; gap: 6px; font: {font}; color: {fore}; background: {back}; padding: 2px 4px;",
                div { style: "width: 13px; height: 13px; border: 1px solid #999; border-radius: 50%; background: white;" }
                "{caption}"
             }
        },
        ControlType::ListBox => rsx! {
             div {
                style: "width: 100%; height: 100%; border: 1px inset #999; padding: 4px; font: {font}; color: {fore}; background: {back}; overflow: auto;",
                {
                    let items = control.get_list_items();
                    if items.is_empty() {
                        rsx! {
                            div { style: "padding: 2px 4px; color: #999;", "(empty)" }
                        }
                    } else {
                        rsx! {
                            for (idx, item) in items.iter().enumerate() {
                                div { 
                                    key: "{idx}",
                                    style: if idx == 0 { "padding: 2px 4px; background: #0078d4; color: white;" } else { "padding: 2px 4px;" },
                                    "{item}"
                                }
                            }
                        }
                    }
                }
             }
        },
        ControlType::ComboBox => rsx! {
             div {
                style: "width: 100%; height: 100%; border: 1px inset #999; display: flex; align-items: center; font: {font}; color: {fore}; background: {back};",
                div { 
                    style: "flex: 1; padding: 2px 4px;", 
                    {
                        let items = control.get_list_items();
                        if items.is_empty() {
                            text.to_string()
                        } else {
                            items.first().cloned().unwrap_or_else(|| text.to_string())
                        }
                    }
                }
                div { 
                    style: "width: 17px; height: 100%; background: #e1e1e1; border-left: 1px solid #999; display: flex; align-items: center; justify-content: center;",
                    "▼"
                }
             }
        },
        ControlType::TreeView => rsx! {
            div {
                style: "width: 100%; height: 100%; background: {back}; border: 1px inset #999; padding: 4px; font: {font}; color: {fore}; overflow: hidden;",
                div { style: "padding: 1px 0; font: {font}; color: {fore};", "▶ Node 1" }
                div { style: "padding: 1px 0; font: {font}; color: {fore};", "▼ Node 2" }
                div { style: "padding: 1px 0 1px 16px; font: {font}; color: {fore};", "▶ Child 1" }
                div { style: "padding: 1px 0 1px 16px; font: {font}; color: {fore};", "  Child 2" }
                div { style: "padding: 1px 0; font: {font}; color: {fore};", "▶ Node 3" }
            }
        },
        ControlType::DataGridView => rsx! {
            div {
                style: "width: 100%; height: 100%; background: white; border: 1px solid #999; font-size: 11px; color: black; overflow: hidden; display: flex; flex-direction: column;",
                // Column headers
                div {
                    style: "display: flex; background: #f0f0f0; border-bottom: 1px solid #ccc; font-weight: bold;",
                    div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #ccc;", "Column1" }
                    div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #ccc;", "Column2" }
                    div { style: "flex: 1; padding: 3px 6px;", "Column3" }
                }
                // Empty rows
                for i in 0..4 {
                    div {
                        key: "{i}",
                        style: "display: flex; border-bottom: 1px solid #eee;",
                        div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #eee; min-height: 20px;" }
                        div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #eee; min-height: 20px;" }
                        div { style: "flex: 1; padding: 3px 6px; min-height: 20px;" }
                    }
                }
            }
        },
        ControlType::ListView => rsx! {
            div {
                style: "width: 100%; height: 100%; background: {back}; border: 1px inset #999; font: {font}; color: {fore}; overflow: hidden; display: flex; flex-direction: column;",
                // Column headers
                div {
                    style: "display: flex; background: #f0f0f0; border-bottom: 1px solid #ccc; font-weight: bold; {font}; color: {fore};",
                    div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #ccc;", "Name" }
                    div { style: "flex: 1; padding: 3px 6px; border-right: 1px solid #ccc;", "Type" }
                    div { style: "flex: 1; padding: 3px 6px;", "Size" }
                }
                div { style: "flex: 1; padding: 4px; color: #999;", "(empty)" }
            }
        },
        ControlType::BindingNavigator => rsx! {
            div {
                style: "width: 100%; height: 100%; display: flex; align-items: center; gap: 2px; background: #f0f0f0; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px; color: black;",
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "⏮" }
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "◀" }
                div { style: "padding: 0 4px; min-width: 30px; text-align: center; border: 1px solid #ccc; background: white;", "0" }
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "▶" }
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "⏭" }
                div { style: "width: 1px; height: 12px; background: #aaa; margin: 0 2px;" }
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "➕" }
                div { style: "padding: 1px 4px; border: 1px solid #aaa; background: #e8e8e8;", "❌" }
            }
        },
        _ => rsx! { div { "..." } }
    }
}

#[component]
fn ResizeHandles(control_bounds: Bounds, drag_state: Signal<Option<DragState>>) -> Element {
    let handle_style = "position: absolute; width: 6px; height: 6px; background: #0078d4; border: 1px solid white;";
    
    let make_handle = move |pos: HandlePosition, top, left, cursor| {
        rsx! {
            div {
                style: "{handle_style} top: {top}; left: {left}; cursor: {cursor};",
                onmousedown: move |evt| {
                    evt.stop_propagation();
                    drag_state.set(Some(DragState {
                        start_x: evt.client_coordinates().x as i32,
                        start_y: evt.client_coordinates().y as i32,
                        initial_bounds: control_bounds,
                        mode: DragMode::Resize(pos)
                    }));
                }
            }
        }
    };
    
    rsx! {
        // Corners
        {make_handle(HandlePosition::TopLeft, "-4px", "-4px", "nw-resize")}
        {make_handle(HandlePosition::TopRight, "-4px", "calc(100% - 4px)", "ne-resize")}
        {make_handle(HandlePosition::BottomLeft, "calc(100% - 4px)", "-4px", "sw-resize")}
        {make_handle(HandlePosition::BottomRight, "calc(100% - 4px)", "calc(100% - 4px)", "se-resize")}
        
        // Sides
        {make_handle(HandlePosition::Top, "-4px", "calc(50% - 4px)", "n-resize")}
        {make_handle(HandlePosition::Bottom, "calc(100% - 4px)", "calc(50% - 4px)", "s-resize")}
        {make_handle(HandlePosition::Left, "calc(50% - 4px)", "-4px", "w-resize")}
        {make_handle(HandlePosition::Right, "calc(50% - 4px)", "calc(100% - 4px)", "e-resize")}
    }
}
