// vybe-test: kotlin/short_circuit/test_iffalse_and_rhs_not_called_even_with_other_predicates
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            fun rhs(): Boolean {
                log += "rhs"
                return true
            }
            __check(((0 > 1) && rhs()).toString(), "false")
            __check((log).toString(), "")
        }
