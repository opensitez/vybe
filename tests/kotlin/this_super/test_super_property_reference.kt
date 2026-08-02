// vybe-test: kotlin/this_super/test_super_property_reference
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open val v = 1 }
class B: A() { override val v = super.v + 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().v).toString(), "3") }
