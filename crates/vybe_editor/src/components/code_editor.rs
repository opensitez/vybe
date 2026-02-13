use dioxus::prelude::*;
use crate::app_state::{AppState, ResourceTarget};
use vybe_forms::EventType;

#[component]
pub fn CodeEditor() -> Element {
    let mut state = use_context::<AppState>();

    // Dropdown state
    let mut selected_object = use_signal(|| "(General)".to_string());
    let mut selected_procedure = use_signal(|| "(Declarations)".to_string());

    // Tab state for VB.NET forms: "code" or "designer"
    let mut code_tab = use_signal(|| "code".to_string());

    // Flag to ignore updates coming from Monaco itself to prevent loops
    let mut self_update = use_signal(|| false);

    // Watch for changes to current_form and update Monaco
    use_effect(move || {
        let current_item = state.current_form.read().clone();
        if current_item.is_some() {
            let code = state.get_current_code();
            let _ = document::eval(&format!(r#"
                if (window.updateMonacoCode) {{
                    window.updateMonacoCode(`{}`);
                }}
            "#, code.replace("`", "\\`").replace("$", "\\$")));
        }
    });

    // Monaco assets are embedded in the binary and served via use_asset_handler.
    // Requests to /assets/vs/... are intercepted and served from embedded data.
    let assets_path = "assets";

    // JS script for Monaco initialization
    let monaco_script = format!(r#"
        const assetsPath = "{}";
        const loaderPath = `${{assetsPath}}/vs/loader.js`;

        function loadScript(src) {{
            return new Promise((resolve, reject) => {{
                if (document.querySelector(`script[src="${{src}}"]`)) {{
                    resolve();
                    return;
                }}
                let script = document.createElement('script');
                script.src = src;
                script.onload = () => resolve();
                script.onerror = (e) => reject(new Error("Script load failed: " + src));
                document.head.appendChild(script);
            }});
        }}

        function createEditor() {{
            // Retry until container is in the DOM
            const container = document.getElementById('monaco-container');
            if (!container) {{
                setTimeout(createEditor, 30);
                return;
            }}

            // Dispose old editor on remount (old container DOM was destroyed)
            if (window.monacoEditor) {{
                window.monacoEditor.dispose();
                window.monacoEditor = null;
            }}

            container.innerText = "";

            window.monacoEditor = monaco.editor.create(container, {{
                value: "",
                language: 'vb',
                theme: 'vs-light',
                automaticLayout: true,
                minimap: {{ enabled: false }}
            }});

            // Change listener
            window.monacoEditor.getModel().onDidChangeContent(() => {{
                const val = window.monacoEditor.getValue();
                dioxus.send({{ type: 'code_change', value: val }});
            }});

            // Global helper to update Monaco content from Rust
            window.updateMonacoCode = function(code) {{
                if (window.monacoEditor) {{
                    const current = window.monacoEditor.getValue();
                    if (current !== code) {{
                        window.monacoEditor.setValue(code);
                    }}
                }}
            }};

            // Helper to jump to a line matching a pattern
            window.jumpToLine = (pattern) => {{
                 const api = window.monacoEditor;
                 if (!api) return;
                 const model = api.getModel();
                 const matches = model.findMatches(pattern, false, false, false, null, true);
                 if (matches.length > 0) {{
                     const range = matches[0].range;
                     api.revealRangeInCenter(range);
                     api.setPosition({{ lineNumber: range.startLineNumber, column: range.startColumn }});
                     api.focus();
                 }}
            }};

            // Signal ready to Rust
            dioxus.send({{ type: 'monaco_ready' }});
        }}

        (async function initMonaco() {{
            try {{
                // If Monaco library already loaded (remount after run/stop), create directly
                if (window.monaco && window.monaco.editor) {{
                    createEditor();
                    return;
                }}

                // First time: load Monaco library
                await loadScript(loaderPath);
                require.config({{ paths: {{ 'vs': `${{assetsPath}}/vs` }} }});
                require(['vs/editor/editor.main'], function() {{
                    createEditor();
                }});
            }} catch (e) {{
                console.error("Failed to load Monaco:", e);
                const el = document.getElementById('monaco-container');
                if(el) el.innerText = "Failed to load Editor: " + (e.message || e);
            }}
        }})();
    "#, assets_path);

    // Track if Monaco is ready
    let mut monaco_ready = use_signal(|| false);

    // use_effect runs AFTER DOM is committed, guaranteeing #monaco-container exists.
    // spawn() starts the async recv loop in the component's scope.
    // On remount (after run/stop), the component is fresh: effect runs again, new eval + spawn.
    use_effect(move || {
        let script = monaco_script.clone();
        let mut handle = document::eval(&script);

        spawn(async move {
            loop {
                match handle.recv::<serde_json::Value>().await {
                    Ok(msg) => {
                        if let Some(obj) = msg.as_object() {
                            if let Some(val_type) = obj.get("type").and_then(|v| v.as_str()) {
                                match val_type {
                                    "monaco_ready" => {
                                        monaco_ready.set(true);
                                        // Sync current state code to Monaco on mount
                                        let code = state.get_current_code();
                                        if !code.is_empty() {
                                            let escaped = code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                            let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                        }
                                    },
                                    "code_change" => {
                                        if let Some(new_code) = obj.get("value").and_then(|v| v.as_str()) {
                                            // Only update if we're in user code tab (not designer read-only)
                                            if *code_tab.read() != "designer" {
                                                self_update.set(true);
                                                state.update_current_code(new_code.to_string());
                                                self_update.set(false);
                                            }
                                        }
                                    },
                                    _ => {}
                                }
                            }
                        }
                    }
                    Err(_) => break,
                }
            }
        });
    });

    // Build object list using a memo so we don't subscribe the component body to project signal
    let objects = use_memo(move || {
        let current_form = state.get_current_form();
        let mut objs = vec!["(General)".to_string()];
        if let Some(form) = &current_form {
            objs.push(form.name.clone());
            for control in &form.controls {
                objs.push(control.name.clone());
            }
        }
        objs
    });

    // Build procedure list based on selected object
    let build_procedures = move |sel_obj: &str| -> Vec<String> {
        if sel_obj == "(General)" {
            vec!["(Declarations)".to_string()]
        } else {
            let mut events = Vec::new();
            let mut c_type = None;
            let current_form = state.get_current_form();
            if let Some(form) = current_form {
                if sel_obj == form.name {
                     c_type = None;
                } else {
                     if let Some(c) = form.get_control_by_name(sel_obj) {
                         c_type = Some(c.control_type);
                     }
                }
            }

            for evt in EventType::all_events() {
                if evt.is_applicable_to(c_type) {
                    events.push(evt.as_str().to_string());
                }
            }
            events
        }
    };

    // Navigation Logic
    let navigate_to_proc = move |obj: String, proc: String| {
        let current_code = state.get_current_code();
        let sub_name = format!("{}_{}", obj, proc);
        let search_pattern = format!("Sub {}", sub_name);

        if current_code.contains(&search_pattern) {
             let _ = document::eval(&format!(r#"
                if (window.jumpToLine) window.jumpToLine("{}");
             "#, search_pattern));
        } else {
             let params = EventType::all_events().iter()
                 .find(|e| e.as_str() == proc)
                 .map(|e| e.parameters())
                 .unwrap_or("");

             let mut new_code = current_code.clone();
             if !new_code.ends_with('\n') {
                 new_code.push('\n');
             }
             new_code.push_str(&format!("\nPrivate Sub {}({})\n\nEnd Sub\n", sub_name, params));
             state.update_current_code(new_code.clone());

             // Sync Monaco with the new code
             let escaped = new_code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
             let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
        }
    };

    rsx! {
        div {
            class: "code-editor",
            style: "flex: 1; display: flex; flex-direction: column; height: 100%;",

            // Top Bar
            div {
                style: "background: #f0f0f0; padding: 4px; border-bottom: 1px solid #ccc; display: flex; gap: 8px;",

                // Object Dropdown
                select {
                    style: "flex: 1;",
                    value: "{selected_object}",
                    onchange: move |evt| {
                        selected_object.set(evt.value());
                        selected_procedure.set("(Declarations)".to_string());
                    },
                    for obj in objects.read().iter() {
                        option { value: "{obj}", "{obj}" }
                    }
                }

                // Procedure Dropdown
                select {
                    style: "flex: 1;",
                    value: "{selected_procedure}",
                    onchange: move |evt| {
                        let proc = evt.value();
                        selected_procedure.set(proc.clone());
                        let obj = selected_object.read().clone();
                        if obj != "(General)" && proc != "(Declarations)" {
                             navigate_to_proc(obj, proc);
                        }
                    },
                    {
                        let current_procs = build_procedures(&selected_object.read());
                        rsx! {
                            option { value: "(Declarations)", "(Declarations)" }
                            if selected_object.read().as_str() != "(General)" {
                                for proc in current_procs {
                                    option { value: "{proc}", "{proc}" }
                                }
                            }
                        }
                    }
                }
            }

            // VB.NET tab bar
            {
                let is_vbnet = state.is_current_form_vbnet();
                rsx! {
                    div {
                        style: "background: #f0f0f0; padding: 2px 4px; border-bottom: 1px solid #eee; font-size: 11px; color: #666; display: flex; align-items: center; gap: 8px;",
                        if is_vbnet {
                            button {
                                style: if *code_tab.read() == "code" { "font-weight: bold; padding: 2px 8px; border: 1px solid #999; background: white;" } else { "padding: 2px 8px; border: 1px solid #ccc; background: #e0e0e0; cursor: pointer;" },
                                onclick: move |_| {
                                    code_tab.set("code".to_string());
                                    // Re-enable editing
                                    let _ = document::eval("if(window.monacoEditor) window.monacoEditor.updateOptions({readOnly: false});");
                                    // Sync Monaco to user code
                                    let code = state.get_current_code();
                                    let escaped = code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                    let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                },
                                "Code"
                            }
                            button {
                                style: if *code_tab.read() == "designer" { "font-weight: bold; padding: 2px 8px; border: 1px solid #999; background: white;" } else { "padding: 2px 8px; border: 1px solid #ccc; background: #e0e0e0; cursor: pointer;" },
                                onclick: move |_| {
                                    code_tab.set("designer".to_string());
                                    // Sync Monaco to designer code (read-only)
                                    let code = state.get_current_designer_code();
                                    let escaped = code.replace('\\', "\\\\").replace('`', "\\`").replace('$', "\\$");
                                    let _ = document::eval(&format!("if(window.updateMonacoCode) window.updateMonacoCode(`{}`);", escaped));
                                    // Make read-only
                                    let _ = document::eval("if(window.monacoEditor) window.monacoEditor.updateOptions({readOnly: true});");
                                },
                                "Designer"
                            }
                            // Show Resources tab if this form has form-level resources
                            {
                                let has_form_resources = state.current_form_has_resources();
                                rsx! {
                                    if has_form_resources {
                                        button {
                                            style: if *code_tab.read() == "resources" { "font-weight: bold; padding: 2px 8px; border: 1px solid #999; background: white;" } else { "padding: 2px 8px; border: 1px solid #ccc; background: #e0e0e0; cursor: pointer;" },
                                            onclick: move |_| {
                                                code_tab.set("resources".to_string());
                                                // Switch to resource view for this form
                                                let form_name = state.current_form.read().clone().unwrap_or_default();
                                                state.show_resources.set(true);
                                                state.show_code_editor.set(true); // stay in code editor panel
                                                state.current_resource_target.set(Some(ResourceTarget::Form(form_name)));
                                            },
                                            "Resources"
                                        }
                                    }
                                }
                            }
                        } else {
                            span { "Code" }
                        }
                    }
                }
            }

            // Monaco Container
            div {
                id: "monaco-container",
                style: "flex: 1; width: 100%; height: 100%; overflow: hidden;",
                "Loading editor..."
            }
        }
    }
}
