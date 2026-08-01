use crate::helpers::run_prints;

#[test]
fn test_float_array_rounding_behaviors() {
    let out = run_prints(
        r#"
        fun main() {
            val a = floatArrayOf(1.2f, 2.7f)
            println(a.sum())
        }
    "#,
    );
    assert_eq!(out, &["3.9"]);
}

#[test]
fn test_float_array_mutation() {
    let out = run_prints(
        r#"
        fun main() {
            val a = FloatArray(2)
            a[0] = 3.5f
            a[1] = a[0] * 2
            println(a[1])
        }
    "#,
    );
    assert_eq!(out, &["7.0"]);
}
