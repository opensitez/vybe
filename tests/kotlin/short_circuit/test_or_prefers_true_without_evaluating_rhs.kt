// vybe-test: kotlin/short_circuit/test_or_prefers_true_without_evaluating_rhs
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
                return false
            }
            __check(((5 > 2) || rhs()).toString(), "true")
            __check((log).toString(), "")
        }
