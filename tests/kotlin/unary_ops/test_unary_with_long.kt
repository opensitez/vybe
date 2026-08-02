// vybe-test: kotlin/unary_ops/test_unary_with_long
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 10L
            __check((-x).toString(), "-10")
            __check((+x).toString(), "10")
        }
