use crate::helpers::run_prints;

#[test]
fn test_string_case_normalization() {
    let out = run_prints(r#"
        fun main() {
            val value = "HeLLo"
            println(value.lowercase())
            println(value.uppercase())
        }
    "#);
    assert_eq!(out, &["hello", "HELLO"]);
}

#[test]
fn test_string_is_blank_checks() {
    let out = run_prints(r#"
        fun main() {
            println("   ".isBlank())
            println("x".isBlank())
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}
