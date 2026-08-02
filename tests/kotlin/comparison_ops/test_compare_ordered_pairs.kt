// vybe-test: kotlin/comparison_ops/test_compare_ordered_pairs
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun cmp(a: Int, b: Int): String {
            return if (a == b) "equal" else if (a < b) "lt" else "gt"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((cmp(1, 1)).toString(), "equal")
            __check((cmp(2, 4)).toString(), "lt")
            __check((cmp(7, 3)).toString(), "gt")
        }
