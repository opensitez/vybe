// vybe-test: kotlin/interfaces/test_interface_reflection_call
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Noting { fun tick(): String }
class Engine : Noting { override fun tick(): String = "ok" }
fun log(n: Noting) { __check((n.tick()).toString(), "ok") }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { log(Engine()) }
