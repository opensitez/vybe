// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_array_present
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val a: IntArray? = intArrayOf(4,5)
        __check((a?.get(1) ?: -1).toString(), "5")
    }
