// vybe-test: kotlin/builtins/test_math_hypot_and_round_trip
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((hypot(3.0, 4.0)).toString(), "5")
            __check((hypot(0.0, 0.0)).toString(), "0")
            __check((sqrt(hypot(3.0, 4.0) * hypot(3.0, 4.0))).toString(), "5")
        }
