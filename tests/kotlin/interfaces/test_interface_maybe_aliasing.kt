// vybe-test: kotlin/interfaces/test_interface_maybe_aliasing
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Named { fun name(): String }
class L : Named { override fun name(): String = "lab" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val n: Named = L()
val m: Named = n
__check((m.name()).toString(), "lab") }
