// vybe-test: kotlin/math_builtins/test_ulp_precision_invariant
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 1.0
            val step = kotlin.math.ulp(v)
            val nearOne = 1.0 + step
            val isAdjacent = kotlin.math.nextAfter(1.0, Double.POSITIVE_INFINITY) == nearOne
            __check((nearOne > v).toString(), "true")
            __check((step > 0.0).toString(), "true")
            __check((isAdjacent).toString(), "true")
            __check((kotlin.math.abs(step - kotlin.math.ulp(nearOne)) < 1e-20).toString(), "true")
        }
