// vybe-test: kotlin/comparison_ops/test_greater_than
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3 > 2).toString(), "true")
            __check((2 > 2).toString(), "false")
            __check((1 > 4).toString(), "false")
        }
