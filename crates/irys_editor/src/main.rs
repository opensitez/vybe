use dioxus::prelude::*;
use irys_project::Project;
use irys_forms::Form;
use dioxus_desktop::{Config, WindowBuilder};
use std::path::PathBuf;

mod app_state;
mod components;

use app_state::AppState;
use components::*;

fn main() {
    // Configure to serve assets from CWD
    // We set resource directory to CWD so 'assets/vs/...' resolves correctly
    let config = Config::new()
        .with_resource_directory(PathBuf::from("."))
        .with_window(
            WindowBuilder::new()
                .with_title("Irys Basic IDE")
                .with_resizable(true)
        );

    LaunchBuilder::desktop()
        .with_cfg(config)
        .launch(App);
}

#[component]
fn App() -> Element {
    // Initialize app state
    use_context_provider(|| {
        let mut state = AppState::new();
        
        // Create default project
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.caption = "Form1".to_string();
        project.add_form(form);
        
        state.project.set(Some(project));
        state.current_form.set(Some("Form1".to_string()));
        
        state
    });
    
    let mut state = use_context::<AppState>();
    let run_mode = *state.run_mode.read();
    let show_toolbox = *state.show_toolbox.read();
    let show_properties = *state.show_properties.read();
    let show_project_explorer = *state.show_project_explorer.read();
    
    rsx! {
        div {
            style: "width: 100vw; height: 100vh; display: flex; flex-direction: column; font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px;",
            
            // Dialogs
            ProjectPropertiesDialog {}
            
            // Menu Bar
            MenuBar {}
            
            // Toolbar
            Toolbar {}
            
            // Main Content Area
            div {
                style: "flex: 1; display: flex; overflow: hidden;",
                
                // Left Sidebar - Project Explorer
                if show_project_explorer {
                    ProjectExplorer {}
                }
                
                // Left Sidebar - Toolbox (only in design mode)
                if !run_mode && show_toolbox {
                    Toolbox {}
                }
                
                // Central Area
                div {
                    style: "flex: 1; display: flex; flex-direction: column;",
                    
                    if run_mode {
                        RuntimePanel {}
                    } else if *state.show_resources.read() {
                        if let Some(proj) = state.project.read().clone() {
                            ResourceEditor {
                                resources: proj.resources.clone(),
                                on_change: move |new_mgr| {
                                    let mut p_lock = state.project.write();
                                    if let Some(p) = p_lock.as_mut() {
                                        p.resources = new_mgr;
                                    }
                                }
                            }
                        }
                    } else if *state.show_code_editor.read() {
                        CodeEditor {}
                    } else {
                        FormDesigner {}
                    }
                }
                
                // Right Sidebar - Properties Panel (only in design mode)
                if !run_mode && show_properties {
                    PropertiesPanel {}
                }
            }
        }
    }
}
