// vybe-test: kotlin/comparison_ops/test_compare_nullable
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Int? = null
            __check((a == null).toString(), "true")
            __check((a != null).toString(), "false")
        }
