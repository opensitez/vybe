use crate::helpers::run_prints;

#[test]
fn test_round_floor_ceil() {
    let out = run_prints(r#"
        import kotlin.math.ceil
        import kotlin.math.floor
        import kotlin.math.round

        fun main() {
            println(round(2.2))
            println(floor(2.8))
            println(ceil(2.2))
        }
    "#);
    assert_eq!(out, &["2.0", "2.0", "3.0"]);
}

#[test]
fn test_rounding_integers_from_floats() {
    let out = run_prints(r#"
        import kotlin.math.round

        fun main() {
            println(round(4.4).toInt())
            println(round(4.6).toInt())
        }
    "#);
    assert_eq!(out, &["4", "5"]);
}
