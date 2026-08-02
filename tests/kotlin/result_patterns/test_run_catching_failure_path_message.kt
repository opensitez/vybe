// vybe-test: kotlin/result_patterns/test_run_catching_failure_path_message
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { "1".toInt() + "a".toInt() }
            __check((value.isFailure).toString(), "true")
            __check((value.exceptionOrNull()?.javaClass?.simpleName).toString(), "NumberFormatException")
        }
