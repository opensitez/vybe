// vybe-test: kotlin/operators/test_custom_contains_with_side_effect
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Gate {
            var probes = 0

            operator fun contains(value: Int): Boolean {
                probes += 1
                return value in 10..20
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
            __check((12 in gate).toString(), "true")
            __check((2 in gate).toString(), "false")
            __check((gate.probes).toString(), "2")
            __check(((5 in gate) || (12 in gate)).toString(), "true")
            __check((gate.probes).toString(), "3")
        }
