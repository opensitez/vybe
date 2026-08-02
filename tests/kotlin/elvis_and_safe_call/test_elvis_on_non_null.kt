// vybe-test: kotlin/elvis_and_safe_call/test_elvis_on_non_null
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: String? = "go"
__check((x ?: "none").toString(), "go") }
