// vybe-test: kotlin/result_patterns/test_result_exception_or_null_for_success_is_null
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { 1 }
            __check((value.exceptionOrNull() == null).toString(), "true")
            __check((value.getOrElse { -1 }).toString(), "1")
        }
