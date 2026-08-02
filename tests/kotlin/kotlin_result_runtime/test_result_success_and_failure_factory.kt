// vybe-test: kotlin/kotlin_result_runtime/test_result_success_and_failure_factory
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok: Result<Int> = Result.success(21)
            val fail: Result<Int> = Result.failure(IllegalStateException("bad"))
            __check((ok.getOrElse { -1 }).toString(), "21")
            __check((fail.getOrElse { it.message?.length ?: 0 }).toString(), "3")
            __check((ok.isSuccess).toString(), "true")
            __check((fail.isFailure).toString(), "true")
        }
