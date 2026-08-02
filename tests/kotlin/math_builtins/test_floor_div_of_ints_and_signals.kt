// vybe-test: kotlin/math_builtins/test_floor_div_of_ints_and_signals
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((-5).floorDiv(2)).toString(), "-3")
            __check(((-5).mod(2)).toString(), "-1")
            __check((5.floorDiv(-2)).toString(), "-3")
            __check((5.mod(-2)).toString(), "1")
        }
