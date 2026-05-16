/// Return the default set of shared .NET namespace imports.
pub fn default_interface_imports() -> Vec<String> {
    vec![
        "system".into(),
        "system.console".into(),
        "system.convert".into(),
        "system.math".into(),
        "system.string".into(),
        "system.array".into(),
        "system.environment".into(),
        // IO
        "system.io".into(),
        "system.io.file".into(),
        "system.io.path".into(),
        "system.io.directory".into(),
        // Collections
        "system.collections".into(),
        "system.collections.generic".into(),
        "system.collections.concurrent".into(),
        // Text
        "system.text".into(),
        "system.text.regularexpressions".into(),
        // Threading
        "system.threading".into(),
        "system.threading.thread".into(),
        "system.threading.tasks".into(),
        // Diagnostics
        "system.diagnostics".into(),
        "system.diagnostics.process".into(),
        "system.diagnostics.stopwatch".into(),
        "system.diagnostics.debug".into(),
        "system.diagnostics.trace".into(),
        // Drawing
        "system.drawing".into(),
        // Net
        "system.net".into(),
        "system.net.sockets".into(),
        // Data
        "system.data".into(),
        "system.data.sqlclient".into(),
        "system.data.oledb".into(),
        // Security
        "system.security.cryptography".into(),
        // XML
        "system.xml.linq".into(),
        // LINQ
        "system.linq".into(),
    ]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_interface_imports_excludes_winforms_specific_entries() {
        let imports = default_interface_imports();
        assert!(imports.contains(&"system".to_string()));
        assert!(imports.contains(&"system.io".to_string()));
        assert!(!imports.contains(&"system.windows.forms".to_string()));
        assert!(!imports.contains(&"application".to_string()));
    }
}