// vybe-test: kotlin/elvis_and_safe_call/test_elvis_chain_multiple
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun first(): String? = null
fun second(): String? = "ok"
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((first() ?: second() ?: "z").toString(), "ok") }
