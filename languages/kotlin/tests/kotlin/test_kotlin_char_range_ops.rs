use crate::helpers::run_prints;

#[test]
fn test_char_range_membership_and_iteration() {
    let out = run_prints(r#"
        fun main() {
            val span = 'b'..'e'
            var text = ""
            for (c in span) { text += c }
            println(text)
            println('a' in span)
            println('d' in span)
        }
    "#);
    assert_eq!(out, &["bcde", "false", "true"]);
}

#[test]
fn test_descending_char_range() {
    let out = run_prints(r#"
        fun main() {
            var out = ""
            for (c in 'e' downTo 'c') {
                out += c
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["edc"]);
}
