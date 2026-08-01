use crate::helpers::run_prints;

#[test]
fn test_smart_cast_with_is_checks() {
    let out = run_prints(
        r#"
        fun describe(value: Any): String {
            return if (value is String) {
                "str:" + value.length
            } else if (value is Int) {
                "int:" + value
            } else {
                "other"
            }
        }

        fun main() {
            println(describe("xy"))
            println(describe(7))
            println(describe(true))
        }
    "#,
    );
    assert_eq!(out, &["str:2", "int:7", "other"]);
}

#[test]
fn test_smart_cast_inside_when_expression() {
    let out = run_prints(
        r#"
        fun score(value: Any): String = when (value) {
            is String -> "s:" + value.length
            is Double -> "d:" + value.toInt()
            is Boolean -> "b:" + value
            else -> "n"
        }

        fun main() {
            println(score("abc"))
            println(score(4.9))
            println(score(false))
        }
    "#,
    );
    assert_eq!(out, &["s:3", "d:4", "b:false"]);
}
