// vybe-test: kotlin/kotlin_result_api/test_result_or_null_default_int
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 8 }
            val bad = runCatching<Int> { throw Exception("n") }
            __check((ok.getOrNull()).toString(), "8")
            __check((bad.getOrNull() ?: 100).toString(), "100")
        }
