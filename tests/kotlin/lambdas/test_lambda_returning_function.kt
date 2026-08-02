// vybe-test: kotlin/lambdas/test_lambda_returning_function
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun makeAdder(offset: Int): (Int) -> Int {
            return { value -> value + offset }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val addFive = makeAdder(5)
            __check((addFive(10)).toString(), "15")
        }
