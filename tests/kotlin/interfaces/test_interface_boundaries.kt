// vybe-test: kotlin/interfaces/test_interface_boundaries
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Counter { fun value(): Int }
class A : Counter { override fun value(): Int = 1 }
class B : Counter { override fun value(): Int = 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((A().value()).toString(), "1")
__check((B().value()).toString(), "2") }
