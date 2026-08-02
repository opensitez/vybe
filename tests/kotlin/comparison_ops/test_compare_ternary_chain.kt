// vybe-test: kotlin/comparison_ops/test_compare_ternary_chain
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = if (1 < 2) if (2 < 3) "yes" else "no" else "no"
            __check((out).toString(), "yes")
        }
