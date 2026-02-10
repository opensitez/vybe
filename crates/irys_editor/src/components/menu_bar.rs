use dioxus::prelude::*;
use crate::app_state::AppState;

#[component]
pub fn MenuBar() -> Element {
    let mut state = use_context::<AppState>();
    let mut active_menu = use_signal(|| None::<String>);
    
    let menu_item_style = "padding: 4px 12px; cursor: pointer; position: relative; user-select: none;";
    let _menu_hover_style = "background: #e0e0e0;";
    let dropdown_style = "
        position: absolute; 
        top: 100%; 
        left: 0; 
        background: white; 
        border: 1px solid #ccc; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.2); 
        min-width: 150px; 
        z-index: 1001;
    ";
    let dropdown_item_style = "padding: 6px 12px; cursor: pointer; &:hover { background: #f0f0f0; }";

    // Helper to close menu
    let mut close_menu = move || active_menu.set(None);

    rsx! {
        div {
            class: "menu-bar",
            style: "display: flex; background: #f0f0f0; border-bottom: 1px solid #ccc; padding: 4px 8px;",
            
            // File Menu
            div {
                style: "{menu_item_style}",
                onclick: move |_| {
                    if *active_menu.read() == Some("File".to_string()) {
                        active_menu.set(None);
                    } else {
                        active_menu.set(Some("File".to_string()));
                    }
                },
                "File"
                if *active_menu.read() == Some("File".to_string()) {
                    div {
                        style: "{dropdown_style}",
                        onclick: move |evt| evt.stop_propagation(), // Prevent closing when clicking dropdown bg
                        
                        div { 
                            style: "{dropdown_item_style}",
                            onclick: move |_| { 
                                state.new_project();
                                close_menu(); 
                            },
                            "New Project" 
                        }
                        div { 
                            style: "{dropdown_item_style}",
                            onclick: move |_| { 
                                state.open_project_dialog();
                                close_menu(); 
                            },
                            "Open Project..." 
                        }
                        div { style: "height: 1px; background: #eee; margin: 2px 0;" }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.save_project();
                                close_menu(); 
                            },
                            "Save Project" 
                        }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.save_project_as();
                                close_menu(); 
                            },
                            "Save Project As..." 
                        }
                        div { style: "height: 1px; background: #eee; margin: 2px 0;" }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| -> () { std::process::exit(0); },
                            "Exit" 
                        }
                    }
                }
            }
            
            // Edit Menu
            div {
                style: "{menu_item_style}",
                onclick: move |_| {
                    if *active_menu.read() == Some("Edit".to_string()) {
                        active_menu.set(None);
                    } else {
                        active_menu.set(Some("Edit".to_string()));
                    }
                },
                "Edit"
                if *active_menu.read() == Some("Edit".to_string()) {
                    div {
                        style: "{dropdown_style}",
                        onclick: move |evt| evt.stop_propagation(),

                        {
                            let has_selection = state.selected_control.read().is_some();
                            let has_clipboard = state.clipboard_control.read().is_some();
                            let disabled_style = "padding: 6px 12px; color: #999; cursor: default;";
                            let enabled_style = dropdown_item_style;
                            rsx! {
                                div {
                                    style: if has_selection { enabled_style } else { disabled_style },
                                    onclick: move |_| {
                                        state.delete_selected_control();
                                        close_menu();
                                    },
                                    "Delete"
                                }
                                div { style: "height: 1px; background: #eee; margin: 2px 0;" }
                                div {
                                    style: if has_selection { enabled_style } else { disabled_style },
                                    onclick: move |_| {
                                        state.cut_selected_control();
                                        close_menu();
                                    },
                                    "Cut"
                                }
                                div {
                                    style: if has_selection { enabled_style } else { disabled_style },
                                    onclick: move |_| {
                                        state.copy_selected_control();
                                        close_menu();
                                    },
                                    "Copy"
                                }
                                div {
                                    style: if has_clipboard { enabled_style } else { disabled_style },
                                    onclick: move |_| {
                                        state.paste_control();
                                        close_menu();
                                    },
                                    "Paste"
                                }
                            }
                        }
                    }
                }
            }

            // Project Menu
            div {
                style: "{menu_item_style}",
                onclick: move |_| {
                     if *active_menu.read() == Some("Project".to_string()) {
                        active_menu.set(None);
                    } else {
                        active_menu.set(Some("Project".to_string()));
                    }
                },
                "Project"
                if *active_menu.read() == Some("Project".to_string()) {
                     div {
                        style: "{dropdown_style}",
                        onclick: move |evt| evt.stop_propagation(),
                        
                        div {
                            style: "{dropdown_item_style}",
                            onclick: move |_| {
                                state.add_new_form();
                                close_menu();
                            },
                            "Add Form"
                        }
                        div {
                            style: "{dropdown_item_style}",
                            onclick: move |_| {
                                state.add_new_vbnet_form();
                                close_menu();
                            },
                            "Add VB.NET Form"
                        }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.show_resources.set(true);
                                state.show_code_editor.set(false);
                                close_menu(); 
                            },
                            "Resources..." 
                        }
                         div { style: "height: 1px; background: #eee; margin: 2px 0;" }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.show_project_properties.set(true); 
                                close_menu(); 
                            },
                            "Project Properties..." 
                        }
                    }
                }
            }
            
             // Run Menu
            div {
                style: "{menu_item_style}",
                onclick: move |_| {
                     if *active_menu.read() == Some("Run".to_string()) {
                        active_menu.set(None);
                    } else {
                        active_menu.set(Some("Run".to_string()));
                    }
                },
                "Run"
                if *active_menu.read() == Some("Run".to_string()) {
                     div {
                        style: "{dropdown_style}",
                        onclick: move |evt| evt.stop_propagation(),
                        
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.run_mode.set(true); 
                                close_menu(); 
                            },
                            "Start" 
                        }
                        div { 
                            style: "{dropdown_item_style}", 
                            onclick: move |_| { 
                                state.run_mode.set(false); 
                                close_menu(); 
                            },
                            "End" 
                        }
                    }
                }
            }
            
             // Close menu when clicking outside (handled by overlay if we had one, 
             // or by global click listener - simplified here by just toggling)
             // For a real app, we'd want a transparent overlay or window listener.
             if active_menu.read().is_some() {
                 div {
                     style: "position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 1000; cursor: default;",
                     onclick: move |_| active_menu.set(None)
                 }
             }
        }
    }
}
