// vybe-test: kotlin/kotlin_result_api/test_result_recover_success_passthrough
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 3 }
            val recovered = ok.recover { -1 }
            __check((recovered.getOrNull()).toString(), "3")
            __check((recovered.isFailure).toString(), "false")
        }
