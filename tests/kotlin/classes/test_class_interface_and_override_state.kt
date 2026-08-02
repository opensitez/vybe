// vybe-test: kotlin/classes/test_class_interface_and_override_state
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class Base { open fun text(): String = "base" }
class Derived: Base() { override fun text(): String = "derived" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b: Base = Derived()
__check((b.text()).toString(), "derived") }
