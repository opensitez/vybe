use crate::helpers::run_prints;

#[test]
fn test_char_classification_methods() {
    let out = run_prints(
        r#"
        fun main() {
            println('A'.isLetter())
            println('3'.isDigit())
            println(' '.isWhitespace())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_char_case_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val c = 'b'
            println(c.toUpperCase())
            println(c.toLowerCase())
        }
    "#,
    );
    assert_eq!(out, &["B", "b"]);
}
