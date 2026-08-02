// vybe-test: kotlin/unary_ops/test_unary_expression_with_nullable
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int? = 5
            __check((value ?: -1).toString(), "5")
            __check((value?.let { -it } ?: -1).toString(), "-5")
        }
