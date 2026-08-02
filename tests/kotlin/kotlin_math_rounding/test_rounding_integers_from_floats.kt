// vybe-test: kotlin/kotlin_math_rounding/test_rounding_integers_from_floats
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_rounding.rs

import kotlin.math.round

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(4.4).toInt()).toString(), "4")
            __check((round(4.6).toInt()).toString(), "5")
        }
