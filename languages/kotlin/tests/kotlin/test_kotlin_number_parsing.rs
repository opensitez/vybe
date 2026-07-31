use crate::helpers::run_prints;

#[test]
fn test_parse_numbers_in_expected_formats() {
    let out = run_prints(r#"
        fun main() {
            println("12".toInt())
            println("3.5".toDouble())
            println("ff".toInt(16))
        }
    "#);
    assert_eq!(out, &["12", "3.5", "255"]);
}

#[test]
fn test_to_int_or_null_handles_invalid_input() {
    let out = run_prints(r#"
        fun main() {
            println("x".toIntOrNull())
            println("".toIntOrNull())
        }
    "#);
    assert_eq!(out, &["null", "null"]);
}
