/// Return the default set of WinForms-specific namespace imports.
pub fn default_interface_imports() -> Vec<String> {
    vec![
        "system.windows.forms".into(),
        // WinForms bare names (for Application.Run, Application.Exit)
        "application".into(),
    ]
}
