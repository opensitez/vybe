// vybe-test: kotlin/builtins/test_rounding_negative_and_positive_edges
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(2.4)).toString(), "2")
            __check((round(-2.6)).toString(), "-3")
            __check((round(-2.0)).toString(), "-2")
        }
