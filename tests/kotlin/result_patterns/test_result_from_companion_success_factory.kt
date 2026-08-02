// vybe-test: kotlin/result_patterns/test_result_from_companion_success_factory
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.success(99)
            __check((value.getOrNull()).toString(), "99")
        }
