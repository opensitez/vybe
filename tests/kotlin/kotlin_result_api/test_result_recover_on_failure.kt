// vybe-test: kotlin/kotlin_result_api/test_result_recover_on_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching<String> { throw IllegalArgumentException("x") }
            val recovered = bad.recover { "fixed" }
            __check((recovered.getOrNull()).toString(), "fixed")
            __check((recovered.isFailure).toString(), "false")
        }
