// vybe-test: kotlin/comparison_ops/test_compare_nested
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1 < 2) == true).toString(), "true")
            __check(((1 < 2) == false).toString(), "false")
        }
