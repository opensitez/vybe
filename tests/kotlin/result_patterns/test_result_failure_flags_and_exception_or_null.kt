// vybe-test: kotlin/result_patterns/test_result_failure_flags_and_exception_or_null
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.failure<Int>(IllegalStateException("bad"))
            __check((value.isSuccess).toString(), "false")
            __check((value.isFailure).toString(), "true")
            __check((value.getOrNull() == null).toString(), "true")
            __check((value.exceptionOrNull() is? Exception).toString(), "true")
        }
