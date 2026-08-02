// vybe-test: kotlin/kotlin_result_api/test_result_get_or_null_for_success
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { "7".toInt() }
            __check((ok.getOrNull()).toString(), "7")
            __check((ok.exceptionOrNull() == null).toString(), "true")
        }
