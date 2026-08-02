// vybe-test: kotlin/advanced_features/test_advanced_inheritance_chain_override
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class A { open fun value(): Int = 1 }
open class B : A() { override fun value(): Int = 2 }
class C : B() { override fun value(): Int = 3 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().value()).toString(), "3") }
