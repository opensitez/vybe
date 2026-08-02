// vybe-test: kotlin/elvis_and_safe_call/test_nonnull_throwing_with_elvis
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x: String? = ""
        __check((x!!.ifEmpty { "empty" }).toString(), "empty")
    }
