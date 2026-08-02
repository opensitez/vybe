// vybe-test: kotlin/elvis_and_safe_call/test_chained_safe_calls
// origin: languages/kotlin/tests/kotlin/test_elvis_and_safe_call.rs

class A { fun child() = B() }
class B { fun name(): String = "b" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a: A? = A()
__check((a?.child()?.name()).toString(), "b") }
