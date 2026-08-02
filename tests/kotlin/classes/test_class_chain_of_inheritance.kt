// vybe-test: kotlin/classes/test_class_chain_of_inheritance
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class A { open fun num(): Int = 1 }
open class B : A() { override fun num(): Int = 2 }
class C : B() { override fun num(): Int = super.num() + 3 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((C().num()).toString(), "5") }
