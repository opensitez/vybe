// vybe-test: kotlin/this_super/test_super_with_override_chain
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open val p = "A" }
open class B: A() { override val p = "B" }
class C: B() { override val p = super<B>.p + "C" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().p).toString(), "BC") }
