// vybe-test: kotlin/short_circuit/test_truth_table_basics_and_or
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true || false).toString(), "true")
            __check((false || false).toString(), "false")
            __check((true && true).toString(), "true")
            __check((false && true).toString(), "false")
        }
