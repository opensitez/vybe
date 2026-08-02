// vybe-test: kotlin/interfaces/test_interface_as_function_argument
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Caller { fun call(): Int }
class Num : Caller { override fun call(): Int = 8 }
fun invoke(c: Caller) = c.call()
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((invoke(Num())).toString(), "8") }
