use std::collections::HashSet;
use std::sync::LazyLock;

static NAMESPACE_ROOTS: LazyLock<HashSet<String>> = LazyLock::new(namespace_roots);

pub fn is_namespace_root(name: &str) -> bool {
    NAMESPACE_ROOTS.contains(name)
}

pub fn namespace_roots() -> HashSet<String> {
    let mut roots = HashSet::new();
    for name in &[
        "system", "microsoft", "vybe",
        "math", "console", "convert", "strings", "array",
        "file", "io", "directory", "environment", "thread", "json", "color",
        "datetime", "stringbuilder", "process", "timespan", "guid", "point", "size",
        "font", "random", "path", "encoding",
        "app", "screen",
        "stopwatch", "debug", "trace",
        "streamreader", "streamwriter", "filestream", "binaryreader", "binarywriter",
        "memorystream", "webrequest", "httpwebrequest", "webclient", "socket",
        "tcpclient", "tcplistener", "udpclient", "task", "timer", "mutex", "semaphore",
        "regex", "match", "list", "dictionary", "queue", "stack", "hashset",
        "concurrentdictionary", "concurrentqueue", "concurrentstack", "concurrentbag",
        "arraylist", "hashtable", "sortedlist", "collection", "datatable", "dataset",
        "datarow", "datacolumn", "sqlconnection", "sqlcommand", "sqldatareader",
        "sqldataadapter", "sqltransaction", "oledbconnection", "oledbcommand",
        "adodb", "connection", "command", "recordset", "xdocument", "xelement", "xmldocument",
        "pen", "solidbrush", "graphics", "bitmap", "image", "colortranslator",
        "systemcolors", "int", "integer", "long", "double", "single", "string",
        "boolean", "byte", "float", "bool", "object", "int32", "int64", "uint32",
    ] {
        roots.insert(name.to_string());
    }
    roots
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_namespace_roots_excludes_winforms_specific_entries() {
        let roots = namespace_roots();
        assert!(roots.contains("console"));
        assert!(roots.contains("graphics"));
        assert!(!roots.contains("application"));
        assert!(!roots.contains("messagebox"));
        assert!(!roots.contains("formborderstyle"));
    }
}