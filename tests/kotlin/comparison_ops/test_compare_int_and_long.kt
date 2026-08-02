// vybe-test: kotlin/comparison_ops/test_compare_int_and_long
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1L > 0).toString(), "true")
            __check((1 == 1L).toString(), "true")
        }
