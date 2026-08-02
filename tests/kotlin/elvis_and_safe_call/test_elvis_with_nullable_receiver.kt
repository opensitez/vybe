// vybe-test: kotlin/elvis_and_safe_call/test_elvis_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val s: String? = null
        __check((s?.uppercase() ?: "NONE").toString(), "NONE")
    }
