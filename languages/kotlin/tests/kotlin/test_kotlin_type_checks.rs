use crate::helpers::run_prints;

#[test]
fn test_is_as_and_as_question_mark_paths() {
    let out = run_prints(
        r#"
        fun main() {
            val values: List<Any?> = listOf("x", 2, null, 3.1)
            val first = values[0] as String
            val second = values[2] as? String
            println(first)
            println(second)
        }
    "#,
    );
    assert_eq!(out, &["x", "null"]);
}

#[test]
fn test_smart_cast_after_type_check() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Any = 12L
            if (value is Long) {
                println(value + 3)
            } else {
                println(0)
            }
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}
