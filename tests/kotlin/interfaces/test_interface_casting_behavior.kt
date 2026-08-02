// vybe-test: kotlin/interfaces/test_interface_casting_behavior
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface X { fun value(): Int }
class Y: X { override fun value(): Int = 4 }
class Z : X { override fun value(): Int = 9 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: X = Y()
__check((x.value() + Z().value()).toString(), "13") }
