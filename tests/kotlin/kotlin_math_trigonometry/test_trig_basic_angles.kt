// vybe-test: kotlin/kotlin_math_trigonometry/test_trig_basic_angles
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_trigonometry.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlin.math.sin(0.0)).toString(), "0.0")
            __check((kotlin.math.cos(0.0)).toString(), "1.0")
        }
