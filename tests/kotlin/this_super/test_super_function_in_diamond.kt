// vybe-test: kotlin/this_super/test_super_function_in_diamond
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun x() = "A" }
interface I { fun x() = "I" }
class B : A(), I { override fun x() = super<A>.x() }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().x()).toString(), "A") }
