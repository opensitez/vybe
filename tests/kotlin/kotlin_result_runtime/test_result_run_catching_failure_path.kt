// vybe-test: kotlin/kotlin_result_runtime/test_result_run_catching_failure_path
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
            __check((result.isFailure).toString(), "true")
            __check((result.isSuccess).toString(), "false")
            __check((result.exceptionOrNull()?.message).toString(), "bad")
            __check((result.getOrNull() == null).toString(), "true")
        }
