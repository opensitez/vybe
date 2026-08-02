// vybe-test: kotlin/comparison_ops/test_compare_sign_inversion
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun isPositive(n: Int): Boolean {
            return n > 0
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isPositive(-1) == !isPositive(1)).toString(), "true")
            __check((!isPositive(0) == isPositive(-1)).toString(), "true")
        }
