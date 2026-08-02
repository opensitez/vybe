// vybe-test: kotlin/this_super/test_super_class_in_inner
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

open class A { open fun label() = "base" }
class B : A() {
    inner class Inner { fun label() = super@B.label() }
}
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((B().Inner().label()).toString(), "base") }
