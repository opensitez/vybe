// vybe-test: kotlin/lambdas/test_trailing_lambda_syntax
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun applyOp(x: Int, op: (Int) -> Int): Int {
            return op(x)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = applyOp(10) { it * 3 }
            __check((result).toString(), "30")
        }
