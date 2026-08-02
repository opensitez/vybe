// vybe-test: kotlin/kotlin_math_rounding/test_round_floor_ceil
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_rounding.rs

import kotlin.math.ceil
        import kotlin.math.floor
        import kotlin.math.round

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((round(2.2)).toString(), "2.0")
            __check((floor(2.8)).toString(), "2.0")
            __check((ceil(2.2)).toString(), "3.0")
        }
