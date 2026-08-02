// vybe-test: kotlin/math_builtins/test_unsigned_power_and_sqrt_chain_edge_case
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val squared = kotlin.math.sqrt(2.0) * kotlin.math.sqrt(2.0)
            val closeToTwo = kotlin.math.abs(squared - 2.0) < 0.0000001
            __check((closeToTwo).toString(), "true")
            __check((kotlin.math.pow(2.0, 0.0)).toString(), "1.0")
            __check((kotlin.math.pow(0.0, 1.0)).toString(), "0.0")
            __check((kotlin.math.pow(0.0, 0.0)).toString(), "1.0")
        }
