// vybe-test: kotlin/boolean_logic/test_boolean_short_circuit_and_side_effects
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var calls = 0
            fun shouldRun() : Boolean {
                calls++
                return true
            }
            __check((false && shouldRun()).toString(), "false")
            __check((calls).toString(), "0")
            __check((true || shouldRun()).toString(), "true")
            __check((calls).toString(), "0")
        }
