// vybe-test: kotlin/builtins/test_abs_and_min_max_boundaries
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((abs(-12)).toString(), "12")
            __check((max(-5, -2)).toString(), "-2")
            __check((min(-5, -2)).toString(), "-5")
        }
