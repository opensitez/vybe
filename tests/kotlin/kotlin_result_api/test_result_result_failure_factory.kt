// vybe-test: kotlin/kotlin_result_api/test_result_result_failure_factory
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Result<String> = Result.failure(IllegalStateException("bad"))
            __check((value.isFailure).toString(), "true")
            __check((value.exceptionOrNull()?.message).toString(), "bad")
        }
