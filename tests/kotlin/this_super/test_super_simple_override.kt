// vybe-test: kotlin/this_super/test_super_simple_override
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun tag() = "a" }
class B: A() { override fun tag() = super.tag() + "+b" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().tag()).toString(), "a+b") }
