// vybe-test: kotlin/kotlin_result_api/test_result_get_or_default
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 9 }
            val bad = runCatching<Int> { throw Exception("bad") }
            __check((ok.getOrDefault(1)).toString(), "9")
            __check((bad.getOrDefault(1)).toString(), "1")
        }
