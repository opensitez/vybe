// vybe-test: kotlin/elvis_and_safe_call/test_elvis_with_expression
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Int? = null
        val y = x?.plus(1) ?: 9
        __check((y).toString(), "9")
    }
