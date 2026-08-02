// vybe-test: kotlin/interfaces/test_interface_object_dispatch
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Op { fun run(x: Int): Int }
class Add : Op { override fun run(x: Int): Int = x + 1 }
class Mul : Op { override fun run(x: Int): Int = x * 2 }
fun apply(op: Op, value: Int): Int = op.run(value)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((apply(Add(), 3)).toString(), "4")
__check((apply(Mul(), 3)).toString(), "6") }
