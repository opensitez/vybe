// vybe-test: kotlin/elvis_and_safe_call/test_elvis_on_nullable_list
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: List<Int>? = null
__check(((x?.size ?: 0) + 1).toString(), "1") }
