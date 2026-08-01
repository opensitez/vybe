use crate::helpers::run_prints;

#[test]
fn test_unit_return_is_executable() {
    let out = run_prints(
        r#"
        var marker = 0

        fun stamp(value: Int): Unit {
            marker = value
        }

        fun main() {
            stamp(7)
            println(marker)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_unit_nullable_slot_exists() {
    let out = run_prints(
        r#"
        fun main() {
            val x: Unit? = null
            println(x == null)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}
