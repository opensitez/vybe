// vybe-test: kotlin/comparison_ops/test_compare_chained
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 2
            __check((x > 1 && x < 5).toString(), "true")
            __check((x < 2 || x > 10).toString(), "false")
        }
