// vybe-test: kotlin/builtins/test_floor_ceil_on_fractional_negative_inputs
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((floor(-3.9)).toString(), "-4")
            __check((ceil(-3.9)).toString(), "-3")
        }
