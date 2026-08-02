// vybe-test: kotlin/lambdas/test_lambda_takes_lambda
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun withOperation(value: Int, transform: (Int) -> Int): Int {
            return transform(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pipeline = withOperation
            __check((pipeline(5, { it + 10 })).toString(), "15")
            __check((pipeline(2, { it * 3 })).toString(), "6")
        }
