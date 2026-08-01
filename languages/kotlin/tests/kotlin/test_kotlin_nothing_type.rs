use crate::helpers::run_prints;

#[test]
fn test_nothing_type_is_used_in_never_returning_flow() {
    let out = run_prints(
        r#"
        fun failNow(): Nothing = throw Exception("x")

        fun main() {
            println(try {
                failNow()
                "bad"
            } catch (e: Exception) {
                "caught"
            })
        }
    "#,
    );
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_nothing_type_in_union_expression() {
    let out = run_prints(
        r#"
        fun boom(): Nothing = throw Exception("nope")

        fun valueOrBoom(v: Int): Int {
            return if (v > 0) v else boom()
        }

        fun main() {
            println(valueOrBoom(2))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}
