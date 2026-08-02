// vybe-test: kotlin/comparison_ops/test_compare_tuples
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

data class P(val x: Int, val y: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((P(1,2) == P(1,2)).toString(), "true")
            __check((P(1,2) == P(2,1)).toString(), "false")
        }
