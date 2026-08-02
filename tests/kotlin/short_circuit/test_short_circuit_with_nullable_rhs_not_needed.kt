// vybe-test: kotlin/short_circuit/test_short_circuit_with_nullable_rhs_not_needed
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = ""
            val text: String? = ""
            fun check(): Boolean {
                log += "check"
                return true
            }
            __check((text != null && check()).toString(), "true")
            __check((log).toString(), "check")
        }
