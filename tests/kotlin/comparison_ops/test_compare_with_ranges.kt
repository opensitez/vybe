// vybe-test: kotlin/comparison_ops/test_compare_with_ranges
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..4
            __check((2 in r).toString(), "true")
            __check((5 in r).toString(), "false")
        }
