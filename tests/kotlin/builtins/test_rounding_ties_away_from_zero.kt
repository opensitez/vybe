// vybe-test: kotlin/builtins/test_rounding_ties_away_from_zero
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(2.5)).toString(), "3")
            __check((round(-2.5)).toString(), "-3")
            __check((round(3.5)).toString(), "4")
        }
