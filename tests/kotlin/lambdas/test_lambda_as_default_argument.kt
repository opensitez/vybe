// vybe-test: kotlin/lambdas/test_lambda_as_default_argument
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun calculate(input: Int, op: (Int) -> Int = { it * 2 }): Int {
            return op(input)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((calculate(4)).toString(), "8")
            __check((calculate(4, { it + 1 })).toString(), "5")
        }
