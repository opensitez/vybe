use dioxus::prelude::*;
use irys_project::Project;
use irys_forms::Form;
use dioxus::desktop::{Config, WindowBuilder, use_asset_handler};
use std::path::PathBuf;
use std::sync::OnceLock;
use include_dir::{include_dir, Dir};
use wry::http::Response;

/// Monaco editor and other assets embedded at compile time.
static EMBEDDED_ASSETS: Dir = include_dir!("$CARGO_MANIFEST_DIR/../../assets");

mod app_state;
mod components;

use app_state::AppState;
use components::*;

/// Holds the project path passed as CLI argument (if any).
static CLI_PROJECT_PATH: OnceLock<Option<PathBuf>> = OnceLock::new();

fn main() {
    // Check for CLI argument: irys_editor [project.vbp|project.vbproj]
    let cli_path = std::env::args().nth(1).map(PathBuf::from);
    CLI_PROJECT_PATH.set(cli_path).ok();

    let config = Config::new()
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
    // Serve embedded Monaco assets from the binary.
    // Requests to /assets/... are intercepted by this handler.
    use_asset_handler("assets", move |request, responder| {
        // The URI path looks like /assets/vs/loader.js — strip the leading /
        let path = request.uri().path().trim_start_matches('/');
        if let Some(file) = EMBEDDED_ASSETS.get_file(path.trim_start_matches("assets/")) {
            let mime = match path.rsplit('.').next().unwrap_or("") {
                "js" => "application/javascript",
                "css" => "text/css",
                "html" => "text/html",
                "json" => "application/json",
                "wasm" => "application/wasm",
                "ttf" => "font/ttf",
                "woff" | "woff2" => "font/woff2",
                "svg" => "image/svg+xml",
                "png" => "image/png",
                _ => "application/octet-stream",
            };
            responder.respond(
                Response::builder()
                    .header("Content-Type", mime)
                    .header("Access-Control-Allow-Origin", "*")
                    .body(file.contents().to_vec())
                    .unwrap(),
            );
        } else {
            responder.respond(
                Response::builder()
                    .status(404)
                    .body(format!("Asset not found: {}", path).into_bytes())
                    .unwrap(),
            );
        }
    });

    // Initialize app state
    use_context_provider(|| {
        let mut state = AppState::new();

        // Check if a project path was passed as CLI argument
        let cli_path = CLI_PROJECT_PATH.get().and_then(|p| p.as_ref());
        if let Some(path) = cli_path {
            eprintln!("[DEBUG] Loading project from CLI argument: {:?}", path);
            match irys_project::load_project_auto(path) {
                Ok(project) => {
                    eprintln!("[DEBUG] CLI project loaded: '{}' with {} forms", project.name, project.forms.len());
                    for f in &project.forms {
                        eprintln!("[DEBUG]   Form '{}': {} controls, text='{}'", f.form.name, f.form.controls.len(), f.form.text);
                        for c in &f.form.controls {
                            eprintln!("[DEBUG]     Control: {} ({:?}) at ({},{}) {}x{}", c.name, c.control_type, c.bounds.x, c.bounds.y, c.bounds.width, c.bounds.height);
                        }
                    }
                    let first_form = project.forms.first().map(|f| f.form.name.clone());
                    state.project.set(Some(project));
                    state.current_form.set(first_form);
                    state.current_project_path.set(Some(path.to_path_buf()));
                    return state;
                }
                Err(e) => {
                    eprintln!("Failed to load project from CLI: {}", e);
                }
            }
        }
        
        // Create default project (fallback) with VB.NET form
        let mut project = Project::new("Project1");
        let mut form = Form::new("Form1");
        form.text = "Form1".to_string();
        form.width = 640;
        form.height = 480;
        
        let designer_code = irys_forms::serialization::designer_codegen::generate_designer_code(&form);
        let user_code = irys_forms::serialization::designer_codegen::generate_user_code_stub("Form1");
        let form_module = irys_project::FormModule::new_vbnet(form, designer_code, user_code);
        project.forms.push(form_module);
        project.startup_object = irys_project::StartupObject::Form("Form1".to_string());
        project.startup_form = Some("Form1".to_string());
        
        state.project.set(Some(project));
        state.current_form.set(Some("Form1".to_string()));
        
        state
    });
    
    let mut state = use_context::<AppState>();
    let run_mode = *state.run_mode.read();
    let show_toolbox = *state.show_toolbox.read();
    let show_properties = *state.show_properties.read();
    let show_project_explorer = *state.show_project_explorer.read();
    let show_resources = *state.show_resources.read();
    let show_code = *state.show_code_editor.read();
    // Only show toolbox/properties when in form designer view (not code, not resources, and viewing a form)
    let in_form_designer = !run_mode && !show_resources && !show_code && state.get_current_form().is_some();
    
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
                
                // Left Sidebar - Toolbox (only in form designer mode)
                if in_form_designer && show_toolbox {
                    Toolbox {}
                }
                
                // Central Area
                div {
                    style: "flex: 1; display: flex; flex-direction: column;",
                    
                    if run_mode {
                        RuntimePanel {}
                    } else if show_resources {
                        {
                            let target = state.current_resource_target.read().clone();
                            let proj_read = state.project.read();
                            if let Some(proj) = proj_read.as_ref() {
                                match &target {
                                    Some(crate::app_state::ResourceTarget::Project(idx)) => {
                                        if let Some(res) = proj.resource_files.get(*idx) {
                                            let res_clone = res.clone();
                                            let idx_copy = *idx;
                                            rsx! {
                                                ResourceEditor {
                                                    resources: res_clone,
                                                    on_change: move |new_mgr: irys_project::ResourceManager| {
                                                        let mut p_lock = state.project.write();
                                                        if let Some(p) = p_lock.as_mut() {
                                                            if let Some(r) = p.resource_files.get_mut(idx_copy) {
                                                                *r = new_mgr;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        } else {
                                            rsx! { div { "Resource file not found" } }
                                        }
                                    }
                                    Some(crate::app_state::ResourceTarget::Form(form_name)) => {
                                        if let Some(fm) = proj.forms.iter().find(|f| &f.form.name == form_name) {
                                            let res_clone = fm.resources.clone();
                                            let fname = form_name.clone();
                                            rsx! {
                                                ResourceEditor {
                                                    resources: res_clone,
                                                    on_change: move |new_mgr: irys_project::ResourceManager| {
                                                        let mut p_lock = state.project.write();
                                                        if let Some(p) = p_lock.as_mut() {
                                                            if let Some(fm) = p.forms.iter_mut().find(|f| f.form.name == fname) {
                                                                fm.resources = new_mgr;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        } else {
                                            rsx! { div { "Form not found" } }
                                        }
                                    }
                                    None => {
                                        rsx! { div { style: "padding: 20px; color: #999;", "Select a resource file from the Project Explorer" } }
                                    }
                                }
                            } else {
                                rsx! {}
                            }
                        }
                    } else if *state.show_code_editor.read() {
                        CodeEditor {}
                    } else {
                        FormDesigner {}
                    }
                }
                
                // Right Sidebar - Properties Panel (only in form designer mode)
                if in_form_designer && show_properties {
                    PropertiesPanel {}
                }
            }
        }
    }
}
