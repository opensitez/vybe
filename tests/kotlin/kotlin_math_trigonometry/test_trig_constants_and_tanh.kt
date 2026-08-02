// vybe-test: kotlin/kotlin_math_trigonometry/test_trig_constants_and_tanh
// origin: languages/kotlin/tests/kotlin/test_kotlin_math_trigonometry.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlin.math.PI).toString(), "3.141592653589793")
            __check((kotlin.math.tanh(0.0)).toString(), "0.0")
        }
