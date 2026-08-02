// vybe-test: kotlin/short_circuit/test_short_circuit_with_nullable_lhs
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text: String? = null
            __check((text != null && text.isNotEmpty()).toString(), "false")
        }
