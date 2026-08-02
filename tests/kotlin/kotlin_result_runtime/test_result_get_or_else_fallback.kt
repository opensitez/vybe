// vybe-test: kotlin/kotlin_result_runtime/test_result_get_or_else_fallback
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 10 }
            val fallback = ok.getOrElse { 0 }
            val bad = runCatching { throw IllegalStateException("x") }
            val fallbackBad = bad.getOrElse { 99 }
            __check((fallback).toString(), "10")
            __check((fallbackBad).toString(), "99")
        }
