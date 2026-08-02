// vybe-test: kotlin/kotlin_result_api/test_result_get_or_else_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching<Int> { throw IllegalStateException("oops") }
            __check((bad.getOrElse { 11 }).toString(), "11")
        }
