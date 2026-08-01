use crate::helpers::run_prints;

#[test]
fn test_if_with_guarded_branching() {
    let out = run_prints(
        r#"
        fun classify(v: Int): String = if (v > 0) {
            "positive"
        } else if (v < 0) {
            "negative"
        } else {
            "zero"
        }

        fun main() {
            println(classify(4))
            println(classify(-1))
            println(classify(0))
        }
    "#,
    );
    assert_eq!(out, &["positive", "negative", "zero"]);
}

#[test]
fn test_when_with_range_and_equality_guards() {
    let out = run_prints(
        r#"
        fun status(code: Int): String = when {
            code in 200..299 -> "ok"
            code in 400..499 -> "client"
            code >= 500 -> "server"
            else -> "other"
        }

        fun main() {
            println(status(204))
            println(status(404))
            println(status(501))
            println(status(101))
        }
    "#,
    );
    assert_eq!(out, &["ok", "client", "server", "other"]);
}
