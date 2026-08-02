// vybe-test: kotlin/comparison_ops/test_compare_float_nans
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 0.0 / 0.0
            __check((x == x).toString(), "false")
            __check((x != x).toString(), "true")
        }
