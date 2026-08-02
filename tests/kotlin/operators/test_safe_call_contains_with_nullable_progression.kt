// vybe-test: kotlin/operators/test_safe_call_contains_with_nullable_progression
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val maybeRange: IntRange? = null
            __check(((maybeRange?.contains(3)) ?: false).toString(), "false")
            val explicit = 1..4
            __check((explicit?.contains(3)).toString(), "true")
        }
