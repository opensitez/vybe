// vybe-test: kotlin/elvis_and_safe_call/test_elvis_numeric_default
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: Int? = null
        __check(((x ?: 2) * 3).toString(), "6")
    }
