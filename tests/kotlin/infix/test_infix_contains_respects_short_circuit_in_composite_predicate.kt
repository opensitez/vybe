// vybe-test: kotlin/infix/test_infix_contains_respects_short_circuit_in_composite_predicate
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Gate {
            var probes = 0
            operator fun contains(value: Int): Boolean {
                probes += 1
                return value % 2 == 0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val gate = Gate()
            val outcome = (1 in gate) || (2 in gate)
            __check((outcome).toString(), "true")
            __check((gate.probes).toString(), "2")
        }
