// vybe-test: kotlin/interfaces/test_interface_with_returning_impl
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Factory { fun make(): Int }
class Maker : Factory { override fun make(): Int = 10 }
fun makeSomething(factory: Factory): Int = factory.make()
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((makeSomething(Maker())).toString(), "10") }
