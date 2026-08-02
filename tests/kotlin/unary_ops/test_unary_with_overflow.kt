// vybe-test: kotlin/unary_ops/test_unary_with_overflow
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = Int.MIN_VALUE
            val y = -x
            __check((y).toString(), "-2147483648")
        }
