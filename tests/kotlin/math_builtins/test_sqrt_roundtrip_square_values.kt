// vybe-test: kotlin/math_builtins/test_sqrt_roundtrip_square_values
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sqrt(144.0)).toString(), "12")
            __check((sqrt(225.0)).toString(), "15")
            __check((sqrt(2.0 * 2.0)).toString(), "2")
        }
