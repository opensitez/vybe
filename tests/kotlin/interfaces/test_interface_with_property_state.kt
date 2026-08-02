// vybe-test: kotlin/interfaces/test_interface_with_property_state
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Counted { var count: Int }
class Counter: Counted { override var count: Int = 0 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val c: Counted = Counter()
c.count = 3
__check((c.count).toString(), "3") }
