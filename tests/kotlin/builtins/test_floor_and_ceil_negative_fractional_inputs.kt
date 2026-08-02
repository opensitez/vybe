// vybe-test: kotlin/builtins/test_floor_and_ceil_negative_fractional_inputs
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((floor(-2.3)).toString(), "-3")
            __check((ceil(-2.3)).toString(), "-2")
        }
