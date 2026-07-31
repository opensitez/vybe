use crate::helpers::run_prints;

#[test]
fn test_throw_caught_and_recovered() {
    let out = run_prints(r#"
        fun main() {
            fun mustBePositive(value: Int): Int {
                if (value <= 0) throw Exception("bad")
                return value
            }

            try {
                println(mustBePositive(2))
            } catch (e: Exception) {
                println("error")
            } finally {
                println("done")
            }
        }
    "#);
    assert_eq!(out, &["2", "done"]);
}

#[test]
fn test_caught_exception_can_be_rethrown() {
    let out = run_prints(r#"
        fun main() {
            try {
                try {
                    throw Exception("inner")
                } catch (e: Exception) {
                    println("inner")
                    throw e
                }
            } catch (e: Exception) {
                println("outer")
            }
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}
