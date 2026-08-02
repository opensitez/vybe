// vybe-test: kotlin/short_circuit/test_bitwise_and_is_not_boolean_short_circuit
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 and 0).toString(), "0")
            __check((1 or 2).toString(), "3")
        }
