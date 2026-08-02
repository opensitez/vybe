// vybe-test: kotlin/builtins/test_pow_fractional_exponent_for_integer_cube_root
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val rooted = round(pow(27.0, 1.0 / 3.0))
            __check((rooted).toString(), "3")
        }
