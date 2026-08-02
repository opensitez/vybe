// vybe-test: kotlin/interfaces/test_interface_default_property
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Identity { val id: Int }
class Item(override val id: Int): Identity
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Item(7).id).toString(), "7") }
