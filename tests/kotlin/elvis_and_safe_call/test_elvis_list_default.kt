// vybe-test: kotlin/elvis_and_safe_call/test_elvis_list_default
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val xs: List<Int>? = null
val out = xs ?: listOf(1,2)
__check((out.size).toString(), "2") }
