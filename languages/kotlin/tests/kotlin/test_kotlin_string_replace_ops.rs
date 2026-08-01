use crate::helpers::run_prints;

#[test]
fn test_string_replace_and_replace_first() {
    let out = run_prints(
        r#"
        fun main() {
            println("aba".replace("a", "x"))
            println("a1b2c".replaceFirst("1", "-"))
        }
    "#,
    );
    assert_eq!(out, &["xbx", "a-b2c"]);
}

#[test]
fn test_string_split_and_joining() {
    let out = run_prints(
        r#"
        fun main() {
            val parts = "a,b,c".split(",")
            println(parts.size)
            println(parts[1])
        }
    "#,
    );
    assert_eq!(out, &["3", "b"]);
}
