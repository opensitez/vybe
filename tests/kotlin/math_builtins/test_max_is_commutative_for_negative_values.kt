// vybe-test: kotlin/math_builtins/test_max_is_commutative_for_negative_values
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(-3, -10)).toString(), "-3")
            __check((max(-10, -3)).toString(), "-3")
            __check((min(-3, -10)).toString(), "-10")
            __check((min(-10, -3)).toString(), "-10")
        }
