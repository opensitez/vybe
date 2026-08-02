// vybe-test: kotlin/kotlin_math_basics/test_math_abs_and_sign
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_basics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlin.math.abs(-12)).toString(), "12")
            __check((kotlin.math.sign(-4.0)).toString(), "-1.0")
        }
