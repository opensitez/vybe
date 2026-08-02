// vybe-test: kotlin/operators/test_short_circuit_avoids_side_effect
// origin: languages/kotlin/tests/kotlin/test_operators.rs

var steps = 0

        fun maybeHappens(flag: Boolean): Boolean {
            steps += 1
            return flag
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((false && maybeHappens(true)).toString(), "false")
            __check((steps).toString(), "0")
            __check((true || maybeHappens(false)).toString(), "true")
            __check((steps).toString(), "0")
        }
