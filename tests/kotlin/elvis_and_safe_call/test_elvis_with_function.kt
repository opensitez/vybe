// vybe-test: kotlin/elvis_and_safe_call/test_elvis_with_function
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun fallback(): String = "fb"
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: String? = null
__check(((x ?: fallback()).length).toString(), "2") }
