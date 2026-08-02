// vybe-test: kotlin/comparison_ops/test_geq
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((2 >= 2).toString(), "true")
            __check((3 >= 2).toString(), "true")
            __check((1 >= 2).toString(), "false")
        }
