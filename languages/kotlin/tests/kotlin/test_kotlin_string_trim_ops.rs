use crate::helpers::run_prints;

#[test]
fn test_string_trim_and_strip_margin() {
    let out = run_prints(r#"
        fun main() {
            println("  ab ".trim())
            println("|a\n|b".trimMargin("|"))
        }
    "#);
    assert_eq!(out, &["ab", "a\nb"]);
}

#[test]
fn test_string_starts_with_and_ends_with() {
    let out = run_prints(r#"
        fun main() {
            val s = "prefix:value"
            println(s.startsWith("pre"))
            println(s.endsWith("value"))
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}
