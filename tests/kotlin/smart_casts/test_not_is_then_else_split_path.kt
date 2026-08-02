// vybe-test: kotlin/smart_casts/test_not_is_then_else_split_path
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 123
            val label = if (value !is String) {
                "not-string"
            } else {
                "is-string"
            }
            __check((label).toString(), "not-string")
        }
