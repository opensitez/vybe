use crate::helpers::run_prints;

#[test]
fn test_simple_string_extension() {
    let out = run_prints(r#"
        fun String.wrap(prefix: String): String = prefix + "-" + this

        fun main() {
            println("core".wrap("pre"))
        }
    "#);
    assert_eq!(out, &["pre-core"]);
}

#[test]
fn test_receiver_extension_with_this_reference() {
    let out = run_prints(r#"
        fun Int.doublePlusOne(): Int = this + this + 1

        fun main() {
            println(3.doublePlusOne())
        }
    "#);
    assert_eq!(out, &["7"]);
}
