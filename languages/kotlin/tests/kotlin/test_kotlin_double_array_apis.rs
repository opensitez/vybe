use crate::helpers::run_prints;

#[test]
fn test_double_array_average() {
    let out = run_prints(r#"
        fun main() {
            val a = doubleArrayOf(2.0, 4.0, 6.0)
            println(a.average())
        }
    "#);
    assert_eq!(out, &["4.0"]);
}

#[test]
fn test_double_array_contains() {
    let out = run_prints(r#"
        fun main() {
            val a = doubleArrayOf(1.1, 2.2, 3.3)
            println(a.contains(2.2))
            println(a.size)
        }
    "#);
    assert_eq!(out, &["true", "3"]);
}
