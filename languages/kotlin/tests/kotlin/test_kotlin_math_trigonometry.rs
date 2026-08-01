use crate::helpers::run_prints;

#[test]
fn test_trig_basic_angles() {
    let out = run_prints(
        r#"
        fun main() {
            println(kotlin.math.sin(0.0))
            println(kotlin.math.cos(0.0))
        }
    "#,
    );
    assert_eq!(out, &["0.0", "1.0"]);
}

#[test]
fn test_trig_constants_and_tanh() {
    let out = run_prints(
        r#"
        fun main() {
            println(kotlin.math.PI)
            println(kotlin.math.tanh(0.0))
        }
    "#,
    );
    assert_eq!(out, &["3.141592653589793", "0.0"]);
}
