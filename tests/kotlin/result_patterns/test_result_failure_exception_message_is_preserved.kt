// vybe-test: kotlin/result_patterns/test_result_failure_exception_message_is_preserved
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { throw Exception("preserve") }
            __check((value.exceptionOrNull()?.message).toString(), "preserve")
        }
