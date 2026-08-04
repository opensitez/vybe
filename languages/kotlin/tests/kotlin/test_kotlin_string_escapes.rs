kotlin_run_cases! {
    test_backslash_escape => (r#"
        fun main() {
            println("path:" + "c\\temp")
            println("quote:" + "\"")
        }
    "#, vec!["path:c\\temp", "quote:\"".into()]),
    test_dollar_and_newline_escaping => (r#"
        fun main() {
            println("dollar:" + "\$")
            println("slash-n:" + "\\n")
        }
    "#, vec!["dollar:$", "slash-n:\\n"]),
    test_tab_escape_text => (r#"
        fun main() {
            println("x" + "\\t" + "y")
            println("a" + "\\r" + "b")
        }
    "#, vec!["x\\ty", "a\\rb"]),
}
