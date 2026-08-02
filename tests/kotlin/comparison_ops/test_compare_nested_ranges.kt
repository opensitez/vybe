// vybe-test: kotlin/comparison_ops/test_compare_nested_ranges
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 5
            val y = x in 1..10
            val z = x !in 10..20
            __check((y).toString(), "true")
            __check((z).toString(), "true")
        }
