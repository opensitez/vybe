// vybe-test: kotlin/elvis_and_safe_call/test_safe_cast_success
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: Any? = 5
val y = x as? Int
__check((y).toString(), "5") }
