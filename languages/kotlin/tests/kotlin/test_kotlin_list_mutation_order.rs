use crate::helpers::run_prints;

#[test]
fn test_mutable_list_basic_updates() {
    let out = run_prints(
        r#"
        fun main() {
            val l = mutableListOf(1, 3, 4)
            l.add(5)
            l[1] = 2
            println(l.toString())
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 4, 5]"]);
}

#[test]
fn test_mutable_list_remove_first() {
    let out = run_prints(
        r#"
        fun main() {
            val l = mutableListOf("x", "y", "z")
            l.removeAt(0)
            println(l.toString())
        }
    "#,
    );
    assert_eq!(out, &["[y, z]"]);
}
