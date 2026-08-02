// vybe-test: kotlin/comparison_ops/test_compare_when
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 4
            val result = when {
                x < 2 -> "small"
                x == 4 -> "four"
                else -> "other"
            }
            __check((result).toString(), "four")
        }
