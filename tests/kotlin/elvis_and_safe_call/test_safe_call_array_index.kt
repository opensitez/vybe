// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_array_index
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val a: IntArray? = null
        __check((a?.get(0) ?: -1).toString(), "-1")
    }
