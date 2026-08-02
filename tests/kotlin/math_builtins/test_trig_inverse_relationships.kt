// vybe-test: kotlin/math_builtins/test_trig_inverse_relationships
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val angle = kotlin.math.asin(kotlin.math.sin(1.0))
            val cosv = kotlin.math.cos(angle)
            val diff = kotlin.math.round((kotlin.math.abs(cosv - kotlin.math.cos(1.0)) * 1e6).toDouble())
            __check((kotlin.math.abs(angle) < 1e-9).toString(), "false")
            __check((diff == 0.0).toString(), "true")
        }
