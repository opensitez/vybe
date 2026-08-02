// vybe-test: kotlin/enums/test_enum_compare_chain
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Signal { RED, GREEN, YELLOW }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = Signal.RED
val b = Signal.GREEN
__check((a == b).toString(), "false")
__check((a != b).toString(), "true") }
