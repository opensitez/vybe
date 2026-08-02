// vybe-test: kotlin/unary_ops/test_unary_negate_zero
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = -0
            __check((x).toString(), "0")
        }
