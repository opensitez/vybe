// vybe-test: kotlin/comparison_ops/test_compare_with_booleans
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            __check((a == b).toString(), "false")
            __check((a != b).toString(), "true")
        }
