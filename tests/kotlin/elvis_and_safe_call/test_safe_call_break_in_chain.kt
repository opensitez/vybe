// vybe-test: kotlin/elvis_and_safe_call/test_safe_call_break_in_chain
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

class N { val v: Int? = null }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val n: N? = N()
__check((n?.v?.plus(1)).toString(), "null") }
