// vybe-test: kotlin/comparison_ops/test_reference_equality_not_eq
// origin: languages/kotlin/tests/kotlin/test_comparison_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Any = "x"
            val b: Any = "x"
            __check((a == b).toString(), "true")
            __check((a === b).toString(), "true")
        }
