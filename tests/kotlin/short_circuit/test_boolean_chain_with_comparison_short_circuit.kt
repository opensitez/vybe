// vybe-test: kotlin/short_circuit/test_boolean_chain_with_comparison_short_circuit
// origin: languages/kotlin/tests/kotlin/test_short_circuit.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var log = 0
            fun left(): Int { log += 1
return 1 }
            fun right(): Int { log += 10
return 2 }
            val v = (left() == 1) && (right() == 3)
            __check((v).toString(), "false")
            __check((log).toString(), "1")
        }
