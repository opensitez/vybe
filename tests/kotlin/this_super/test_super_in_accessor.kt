// vybe-test: kotlin/this_super/test_super_in_accessor
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun get(): Int = 1 }
class B : A() { override fun get() = super.get() + 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().get()).toString(), "3") }
