// vybe-test: kotlin/math_builtins/test_hypot_uses_pythagorean_contract
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = kotlin.math.hypot(3.0, 4.0)
            __check((h).toString(), "5.0")
            __check((hypot(5.0, 12.0)).toString(), "13.0")
        }
