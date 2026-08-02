// vybe-test: kotlin/interfaces/test_interface_property_polymorphism
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Flag { val code: Int }
class True: Flag { override val code = 1 }
class False: Flag { override val code = 0 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val f: Flag = True()
val g: Flag = False()
__check((f.code + g.code).toString(), "1") }
