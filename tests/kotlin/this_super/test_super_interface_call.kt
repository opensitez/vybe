// vybe-test: kotlin/this_super/test_super_interface_call
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

interface I { fun y() = "I" }
class C : I { override fun y() = super<I>.y() }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().y()).toString(), "I") }
