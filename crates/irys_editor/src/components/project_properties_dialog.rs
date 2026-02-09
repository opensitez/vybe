use dioxus::prelude::*;
use crate::app_state::AppState;

#[component]
pub fn ProjectPropertiesDialog() -> Element {
    let mut state = use_context::<AppState>();
    let show = *state.show_project_properties.read();
    
    if !show {
        return rsx! {};
    }
    
    // Get project name
    let project_name = if let Some(proj) = state.project.read().as_ref() {
        proj.name.clone()
    } else {
        "Project".to_string()
    };
    
    rsx! {
        div {
            class: "modal-overlay",
            style: "
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                background: rgba(0, 0, 0, 0.5);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 1000;
            ",
            
            div {
                class: "modal-content",
                style: "
                    background: white;
                    width: 400px;
                    border: 1px solid #999;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
                    display: flex;
                    flex-direction: column;
                ",
                
                // Header
                div {
                    style: "
                        background: linear-gradient(to bottom, #0078d4, #005a9e);
                        color: white;
                        padding: 6px 10px;
                        font-weight: bold;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    ",
                    span { "{project_name} - Project Properties" }
                    div {
                        style: "cursor: pointer; font-family: monospace; font-weight: bold;",
                        onclick: move |_| state.show_project_properties.set(false),
                        "X"
                    }
                }
                
                // Content
                div {
                    style: "padding: 16px; flex: 1;",
                    
                    div {
                        style: "margin-bottom: 12px;",
                        label { style: "display: block; margin-bottom: 4px; font-weight: bold;", "Project Name:" }
                        input {
                            style: "width: 100%; padding: 4px; border: 1px solid #ccc;",
                            value: "{project_name}",
                            // Read-only for now or implement rename logic via project.write()
                            readonly: true
                        }
                    }
                    
                    div {
                        style: "margin-bottom: 12px;",
                        label { style: "display: block; margin-bottom: 4px; font-weight: bold;", "Startup Object:" }
                        {
                            if let Some(proj) = state.project.read().as_ref() {
                                let current_startup = proj.startup_form.clone();
                                rsx! {
                                    select {
                                        style: "width: 100%; padding: 4px; border: 1px solid #ccc;",
                                        value: "{current_startup.clone().unwrap_or_default()}",
                                        onchange: move |evt| {
                                            let selected_form = evt.value();
                                            if let Some(proj) = state.project.write().as_mut() {
                                                if selected_form.is_empty() {
                                                    proj.startup_form = None;
                                                } else {
                                                    proj.startup_form = Some(selected_form);
                                                }
                                            }
                                        },
                                        option {
                                            value: "",
                                            selected: current_startup.is_none(),
                                            "(None)"
                                        }
                                        for form_module in &proj.forms {
                                            {
                                                let form_name = form_module.form.name.clone();
                                                let is_selected = current_startup.as_ref() == Some(&form_name);
                                                rsx! {
                                                    option {
                                                        key: "{form_name}",
                                                        value: "{form_name}",
                                                        selected: is_selected,
                                                        "{form_name}"
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                rsx! {
                                    select {
                                        style: "width: 100%; padding: 4px; border: 1px solid #ccc;",
                                        disabled: true,
                                        option { "(No project loaded)" }
                                    }
                                }
                            }
                        }
                    }
                    
                    div {
                         style: "font-size: 12px; color: #666; margin-top: 20px;",
                         "More settings coming soon..."
                    }
                }
                
                // Footer
                div {
                    style: "
                        padding: 10px;
                        border-top: 1px solid #ccc;
                        background: #f0f0f0;
                        display: flex;
                        justify-content: flex-end;
                        gap: 8px;
                    ",
                    
                    button {
                        style: "padding: 4px 16px; min-width: 70px;",
                        onclick: move |_| state.show_project_properties.set(false),
                        "OK"
                    }
                    
                    button {
                        style: "padding: 4px 16px; min-width: 70px;",
                        onclick: move |_| state.show_project_properties.set(false),
                        "Cancel"
                    }
                }
            }
        }
    }
}
