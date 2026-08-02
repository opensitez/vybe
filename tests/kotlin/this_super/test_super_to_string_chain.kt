// vybe-test: kotlin/this_super/test_super_to_string_chain
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { override fun toString() = "A" }
class B : A() { override fun toString() = super.toString() + "B" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().toString()).toString(), "AB") }
