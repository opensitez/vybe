use crate::helpers::run_prints;

#[test]
fn test_set_duplicate_elimination() {
    let out = run_prints(r#"
        fun main() {
            val s = setOf(1, 1, 2, 3)
            println(s.size)
            println(s.contains(2))
        }
    "#);
    assert_eq!(out, &["3", "true"]);
}

#[test]
fn test_set_additional_distinct_elements() {
    let out = run_prints(r#"
        fun main() {
            val s = mutableSetOf("a")
            s.add("b")
            s.add("a")
            println(s.size)
            println(s.contains("b"))
        }
    "#);
    assert_eq!(out, &["2", "true"]);
}
