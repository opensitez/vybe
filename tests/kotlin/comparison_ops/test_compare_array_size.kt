// vybe-test: kotlin/comparison_ops/test_compare_array_size
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(1,2,3)
            val b = intArrayOf(1,2)
            __check((a.size > b.size).toString(), "true")
            __check((a.size == b.size).toString(), "false")
        }
