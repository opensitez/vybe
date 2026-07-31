use crate::helpers::run_prints;

#[test]
fn test_math_abs_and_sign() {
    let out = run_prints(r#"
        fun main() {
            println(kotlin.math.abs(-12))
            println(kotlin.math.sign(-4.0))
        }
    "#);
    assert_eq!(out, &["12", "-1.0"]);
}

#[test]
fn test_math_pow_and_sqrt() {
    let out = run_prints(r#"
        fun main() {
            println(kotlin.math.sqrt(81.0))
            println(kotlin.math.pow(2.0, 3.0))
        }
    "#);
    assert_eq!(out, &["9.0", "8.0"]);
}
