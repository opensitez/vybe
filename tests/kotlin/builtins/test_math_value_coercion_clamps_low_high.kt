// vybe-test: kotlin/builtins/test_math_value_coercion_clamps_low_high
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((5.coerceIn(1, 3)).toString(), "3")
            __check(((-1).coerceAtLeast(0)).toString(), "0")
            __check((10.coerceAtMost(7)).toString(), "7")
            __check((4.coerceIn(1, 4)).toString(), "4")
        }
