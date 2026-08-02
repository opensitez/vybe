// vybe-test: kotlin/kotlin_result_api/test_result_get_or_else_success
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 5 }
            __check((ok.getOrElse { 0 }).toString(), "5")
        }
