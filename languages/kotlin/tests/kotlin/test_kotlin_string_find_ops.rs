use crate::helpers::run_prints;

#[test]
fn test_string_search_apis() {
    let out = run_prints(
        r#"
        fun main() {
            val s = "banana"
            println(s.indexOf("na"))
            println(s.lastIndexOf("na"))
            println(s.contains("an"))
        }
    "#,
    );
    assert_eq!(out, &["2", "4", "true"]);
}

#[test]
fn test_string_substring_edges() {
    let out = run_prints(
        r#"
        fun main() {
            val s = "abcdef"
            println(s.substring(1, 3))
            println(s.substring(2))
        }
    "#,
    );
    assert_eq!(out, &["bc", "cdef"]);
}
