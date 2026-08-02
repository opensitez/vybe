// vybe-test: kotlin/elvis_and_safe_call/test_elvis_with_method_call
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val s: String? = "ok"
        __check((s?.uppercase()?.substring(0, 1) ?: "missing").toString(), "O")
    }
