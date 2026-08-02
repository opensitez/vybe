// vybe-test: kotlin/elvis_and_safe_call/test_nullable_return_chain
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun find(v: Boolean): String? = if (v) "yes" else null
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val out = find(false) ?: find(true) ?: "none"
__check((out).toString(), "yes") }
