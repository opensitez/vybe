use crate::helpers::run_prints;

#[test]
fn test_forced_cast_to_expected_type() {
    let out = run_prints(
        r#"
        fun main() {
            val any: Any = 42
            val value = any as Int
            println(value + 1)
        }
    "#,
    );
    assert_eq!(out, &["43"]);
}

#[test]
fn test_nullable_smart_safe_cast() {
    let out = run_prints(
        r#"
        fun main() {
            val maybe: Any? = null
            println((maybe as? String))
            println(("ok" as? String))
        }
    "#,
    );
    assert_eq!(out, &["null", "ok"]);
}
