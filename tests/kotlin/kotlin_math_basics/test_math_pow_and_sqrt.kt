// vybe-test: kotlin/kotlin_math_basics/test_math_pow_and_sqrt
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_basics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlin.math.sqrt(81.0)).toString(), "9.0")
            __check((kotlin.math.pow(2.0, 3.0)).toString(), "8.0")
        }
