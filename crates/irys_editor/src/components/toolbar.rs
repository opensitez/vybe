use dioxus::prelude::*;
use crate::app_state::AppState;

#[component]
pub fn Toolbar() -> Element {
    let mut state = use_context::<AppState>();
    let run_mode = *state.run_mode.read();

    rsx! {
        div {
            class: "toolbar",
            style: "
                height: 36px;
                background: #f0f0f0;
                border-bottom: 1px solid #ccc;
                display: flex;
                align-items: center;
                padding: 0 8px;
                gap: 8px;
            ",
            
            // Run Button
            div {
                style: {
                    let bg = if run_mode { "#e0e0e0" } else { "transparent" };
                    let op = if run_mode { "0.5" } else { "1.0" };
                    let pe = if run_mode { "none" } else { "auto" };
                    format!("padding: 4px; border: 1px solid transparent; border-radius: 3px; cursor: pointer; background: {}; opacity: {}; pointer-events: {};", bg, op, pe)
                },
                title: "Start",
                onclick: move |_| {
                    if !run_mode {
                        state.run_mode.set(true);
                    }
                },
                // Play Icon (Triangle)
                svg {
                    width: "16",
                    height: "16",
                    view_box: "0 0 16 16",
                    path {
                        d: "M4 2 L14 8 L4 14 Z",
                        fill: "#0078d4",
                        stroke: "#005a9e",
                        stroke_width: "1"
                    }
                }
            }
            
            // Stop Button
            div {
                style: {
                    let opacity = if !run_mode { "0.5" } else { "1.0" };
                    let pe = if !run_mode { "none" } else { "auto" };
                    format!("padding: 4px; border: 1px solid transparent; border-radius: 3px; cursor: pointer; opacity: {}; pointer-events: {};", opacity, pe)
                },
                title: "End",
                onclick: move |_| {
                    if run_mode {
                        state.run_mode.set(false);
                    }
                },
                // Stop Icon (Square)
                svg {
                    width: "16",
                    height: "16",
                    view_box: "0 0 16 16",
                    rect {
                        x: "3",
                        y: "3",
                        width: "10",
                        height: "10",
                        fill: "#d13438",
                        stroke: "#a4262c",
                        stroke_width: "1"
                    }
                }
            }
            
            div { style: "width: 1px; height: 20px; background: #ccc; margin: 0 4px;" }
            
            // Toggle Code View Button
            div {
                style: "
                    padding: 4px;
                    border: 1px solid transparent;
                    border-radius: 3px;
                    cursor: pointer;
                ",
                title: "View Code",
                onclick: move |_| state.show_code_editor.set(true),
                // Code Icon
                svg {
                    width: "16",
                    height: "16",
                    view_box: "0 0 16 16",
                    path {
                        d: "M10 2 L8 2 L3 14 L5 14 Z M13.5 6 L16 8 L13.5 10 M2.5 6 L0 8 L2.5 10",
                        stroke: "black",
                        fill: "none",
                        stroke_width: "1.5"
                    }
                }
            }

            // Toggle Object View Button
             div {
                style: "
                    padding: 4px;
                    border: 1px solid transparent;
                    border-radius: 3px;
                    cursor: pointer;
                ",
                title: "View Object",
                onclick: move |_| state.show_code_editor.set(false),
                // Form Icon
                svg {
                    width: "16",
                    height: "16",
                    view_box: "0 0 16 16",
                    rect {
                        x: "2",
                        y: "2",
                        width: "12",
                        height: "11",
                        stroke: "black",
                        fill: "none",
                        stroke_width: "1.5"
                    }
                    rect { x: "2", y: "2", width: "12", height: "3", fill: "black" }
                }
            }
            
            // Toggle Resources Button
             div {
                style: {
                    let is_active = *state.show_resources.read();
                    let bg = if is_active { "#e0e0e0" } else { "transparent" };
                    let border = if is_active { "#bbb" } else { "transparent" };
                    format!("
                        padding: 4px;
                        border: 1px solid {border};
                        background: {bg};
                        border-radius: 3px;
                        cursor: pointer;
                    ")
                },
                title: "View Resources",
                onclick: move |_| {
                    state.show_resources.set(true);
                    state.show_code_editor.set(false);
                },
                // Resource Icon (Table-like)
                svg {
                    width: "16",
                    height: "16",
                    view_box: "0 0 16 16",
                    path {
                        d: "M2 3 L14 3 L14 13 L2 13 Z M2 6 L14 6 M6 3 L6 13",
                        stroke: "black",
                        fill: "none",
                        stroke_width: "1.5"
                    }
                }
            }
            
            div { style: "width: 1px; height: 20px; background: #ccc; margin: 0 4px;" }
            
            // Add Form Button
            button {
                style: "padding: 4px 8px; font-size: 12px; cursor: pointer; border: 1px solid #999; background: white;",
                title: "Add Form",
                onclick: move |_| state.add_new_form(),
                "📋 Form"
            }
            
            // Add Code File Button
            button {
                style: "padding: 4px 8px; font-size: 12px; cursor: pointer; border: 1px solid #999; background: white;",
                title: "Add Code File",
                onclick: move |_| state.add_code_file(),
                "\u{1F4C4} Code"
            }
        }
    }
}
