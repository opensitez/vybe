// vybe-test: kotlin/kotlin_result_runtime/test_result_run_catching_success_path
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching { 5 + 7 }
            __check((result.isSuccess).toString(), "true")
            __check((result.getOrNull()).toString(), "12")
            __check((result.exceptionOrNull() == null).toString(), "true")
        }
