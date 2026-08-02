// vybe-test: kotlin/lambdas/test_lambda_as_argument_list_value
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun apply(x: Int, y: Int, op: (Int, Int) -> Int): Int {
            return op(x, y)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(8, 4, { a, b -> a - b })).toString(), "4")
            __check((apply(2, 2, { a, b -> a / b })).toString(), "1")
        }
