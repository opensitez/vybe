// vybe-test: kotlin/this_super/test_this_in_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun who() = "A" }
open class B: A() { override fun who() = "B" }
class C: B() { override fun who() = super<B>.who() + "C" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().who()).toString(), "BC") }
