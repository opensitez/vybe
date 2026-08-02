// vybe-test: kotlin/this_super/test_super_in_init_order
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open val a = "a" }
class B : A() { override val a = super.a + "+b" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().a).toString(), "a+b") }
