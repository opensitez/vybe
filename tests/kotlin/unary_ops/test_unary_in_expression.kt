// vybe-test: kotlin/unary_ops/test_unary_in_expression
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun sign(v: Int): Int {
            return if (v > 0) +1 else -1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sign(3)).toString(), "1")
            __check((sign(-2)).toString(), "-1")
        }
