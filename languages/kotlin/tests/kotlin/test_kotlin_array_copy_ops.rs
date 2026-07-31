use crate::helpers::run_prints;

#[test]
fn test_array_copy_to_new_reference() {
    let out = run_prints(r#"
        fun main() {
            val a = arrayOf(1, 2, 3)
            val b = a.copyOf()
            b[0] = 9
            println(a[0])
            println(b[0])
        }
    "#);
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_array_copy_of_range() {
    let out = run_prints(r#"
        fun main() {
            val a = arrayOf("a", "b", "c", "d")
            val b = a.copyOfRange(1, 3)
            println(b.toString())
        }
    "#);
    assert_eq!(out, &["[b, c]"]);
}
