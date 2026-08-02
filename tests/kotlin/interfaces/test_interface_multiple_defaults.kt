// vybe-test: kotlin/interfaces/test_interface_multiple_defaults
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface A { fun x(): String = "a" }
interface B { fun y(): String = "b" }
class C : A, B
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b: B = C()
__check((b.y()).toString(), "b") }
