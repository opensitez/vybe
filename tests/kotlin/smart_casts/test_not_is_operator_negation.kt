// vybe-test: kotlin/smart_casts/test_not_is_operator_negation
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = 42
            __check((value !is String).toString(), "true")
        }
