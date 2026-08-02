// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_in_string_concat
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val s: String? = null
        __check(((s ?: "") + "x").toString(), "x")
    }
