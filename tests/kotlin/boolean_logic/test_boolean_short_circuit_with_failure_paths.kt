// vybe-test: kotlin/boolean_logic/test_boolean_short_circuit_with_failure_paths
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var calls = 0
            fun sideEffect() : Boolean {
                calls++
                throw Error("boom")
            }
            __check((false && sideEffect()).toString(), "false")
            __check((calls).toString(), "0")
            __check((true || sideEffect()).toString(), "true")
            __check((calls).toString(), "0")
        }
