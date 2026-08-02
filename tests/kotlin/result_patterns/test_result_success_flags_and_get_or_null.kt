// vybe-test: kotlin/result_patterns/test_result_success_flags_and_get_or_null
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.success(7)
            __check((value.isSuccess).toString(), "true")
            __check((value.isFailure).toString(), "false")
            __check((value.getOrNull()).toString(), "7")
            __check((value.exceptionOrNull() == null).toString(), "true")
        }
