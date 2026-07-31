use crate::helpers::run_prints;

#[test]
fn test_int_array_sum() {
    let out = run_prints(r#"
        fun main() {
            val a = intArrayOf(5, 6, 7)
            var s = 0
            for (v in a) { s += v }
            println(s)
        }
    "#);
    assert_eq!(out, &["18"]);
}

#[test]
fn test_int_array_setter_and_getter() {
    let out = run_prints(r#"
        fun main() {
            val a = IntArray(3)
            a[0] = 4
            a[1] = 5
            a[2] = a[0] + a[1]
            println(a[2])
        }
    "#);
    assert_eq!(out, &["9"]);
}
