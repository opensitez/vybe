// vybe-test: kotlin/kotlin_result_api/test_result_is_success_and_failure_flags
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 10 }
            val bad = runCatching<Int> { throw IllegalArgumentException("bad") }
            __check((ok.isSuccess).toString(), "true")
            __check((ok.isFailure).toString(), "false")
            __check((bad.isSuccess).toString(), "false")
            __check((bad.isFailure).toString(), "true")
        }
