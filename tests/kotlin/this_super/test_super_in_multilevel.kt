// vybe-test: kotlin/this_super/test_super_in_multilevel
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun depth() = "A" }
open class B : A() { override fun depth() = super<A>.depth() + "->B" }
class C : B() { override fun depth() = super<B>.depth() + "->C" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().depth()).toString(), "A->B->C") }
