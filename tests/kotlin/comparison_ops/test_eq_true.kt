// vybe-test: kotlin/comparison_ops/test_eq_true
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 == 1).toString(), "true")
            __check((1 == 2).toString(), "false")
        }
