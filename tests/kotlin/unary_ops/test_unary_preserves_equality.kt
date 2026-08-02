// vybe-test: kotlin/unary_ops/test_unary_preserves_equality
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = -(-3)
            val b = 3
            __check((a == b).toString(), "true")
        }
