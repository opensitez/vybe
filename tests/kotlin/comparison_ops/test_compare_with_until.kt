// vybe-test: kotlin/comparison_ops/test_compare_with_until
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1 until 4
            __check((3 in r).toString(), "true")
            __check((4 in r).toString(), "false")
            __check((0 in r).toString(), "false")
        }
