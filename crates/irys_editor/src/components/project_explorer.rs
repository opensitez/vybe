use dioxus::prelude::*;
use crate::app_state::AppState;

#[component]
pub fn ProjectExplorer() -> Element {
    let mut state = use_context::<AppState>();
    let project = state.project.read();
    let mut current_form = state.current_form;
    
    rsx! {
        div {
            class: "project-explorer",
            style: "width: 200px; background: #fafafa; border-right: 1px solid #ccc; padding: 8px;",
            
            h3 { style: "margin: 0 0 8px 0; font-size: 14px;", "Project Explorer" }
            
            // Toolbar
            div {
                style: "display: flex; gap: 4px; margin-bottom: 8px; border-bottom: 1px solid #ccc; padding-bottom: 4px;",
                button {
                    style: "padding: 2px 6px; font-size: 11px; cursor: pointer;",
                    onclick: move |_| state.show_code_editor.set(false),
                    "View Object"
                }
                button {
                    style: "padding: 2px 6px; font-size: 11px; cursor: pointer;",
                    onclick: move |_| state.show_code_editor.set(true),
                    "View Code"
                }
            }
            
            div {
                style: "border-top: 1px solid #ccc; padding-top: 8px;",
                
                {
                    if let Some(proj) = project.as_ref() {
                        rsx! {
                            div {
                                style: "font-weight: bold; margin-bottom: 8px;",
                                "📁 {proj.name}"
                            }
                            
                            div {
                                style: "margin-left: 16px;",
                                
                                div {
                                    style: "font-weight: bold; margin-bottom: 4px;",
                                    "📋 Forms"
                                }
                                
                                for form_module in &proj.forms {
                                    {
                                        let form_name = form_module.form.name.clone();
                                        let is_selected = *current_form.read() == Some(form_name.clone());
                                        let bg_color = if is_selected { "#e3f2fd" } else { "transparent" };
                                        
                                        rsx! {
                                            div {
                                                key: "{form_name}",
                                                style: "padding: 4px 8px; cursor: pointer; background: {bg_color}; border-radius: 3px; margin-bottom: 2px;",
                                                onclick: move |_| {
                                                    current_form.set(Some(form_name.clone()));
                                                    state.show_code_editor.set(false); // Switch to designer view
                                                },
                                                "  {form_name}"
                                            }
                                        }
                                    }
                                }
                                
                                // Code files section
                                if !proj.code_files.is_empty() {
                                    div {
                                        style: "font-weight: bold; margin-top: 12px; margin-bottom: 4px;",
                                        "\u{1F4C4} Code"
                                    }
                                    for code_file in &proj.code_files {
                                        {
                                            let cf_name = code_file.name.clone();
                                            let is_selected = *current_form.read() == Some(cf_name.clone());
                                            let bg_color = if is_selected { "#e3f2fd" } else { "transparent" };

                                            rsx! {
                                                div {
                                                    key: "{cf_name}",
                                                    style: "padding: 4px 8px; cursor: pointer; background: {bg_color}; border-radius: 3px; margin-bottom: 2px;",
                                                    onclick: move |_| {
                                                        current_form.set(Some(cf_name.clone()));
                                                        state.show_code_editor.set(true);
                                                    },
                                                    "  {cf_name}"
                                                }
                                            }
                                        }
                                    }
                                }

                                // Project Resources
                                {
                                    let is_res_selected = *state.show_resources.read();
                                    let bg_color = if is_res_selected { "#e3f2fd" } else { "transparent" };
                                    rsx! {
                                        div {
                                            style: "margin-top: 12px; font-weight: bold; margin-bottom: 4px;",
                                            "⚙️ Config"
                                        }
                                        div {
                                            style: "padding: 4px 8px; cursor: pointer; background: {bg_color}; border-radius: 3px; margin-bottom: 2px;",
                                            onclick: move |_| {
                                                state.show_resources.set(true);
                                                state.show_code_editor.set(false);
                                            },
                                            "  Project Resources"
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        rsx! {
                            div {
                                style: "color: #999; font-style: italic;",
                                "No project loaded"
                            }
                        }
                    }
                }
            }
        }
    }
}
