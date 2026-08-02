// vybe-test: kotlin/interfaces/test_interface_nested_implementation
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Reader { fun read(): String }
class Bundle { fun create(): Reader = object : Reader { override fun read(): String = "yes" } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b = Bundle()
__check((b.create().read()).toString(), "yes") }
