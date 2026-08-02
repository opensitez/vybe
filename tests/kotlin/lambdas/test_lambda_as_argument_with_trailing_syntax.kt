// vybe-test: kotlin/lambdas/test_lambda_as_argument_with_trailing_syntax
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun applyTwice(value: Int, op: (Int) -> Int): Int {
            return op(op(value))
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = applyTwice(3) { it * 2 }
            __check((result).toString(), "12")
        }
