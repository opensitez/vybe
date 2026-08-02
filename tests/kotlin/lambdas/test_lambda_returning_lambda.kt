// vybe-test: kotlin/lambdas/test_lambda_returning_lambda
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun makeMultiplier(scale: Int): (Int) -> Int {
            return { x: Int -> x * scale }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val times = makeMultiplier(4)
            val plus = makeMultiplier(1)
            __check((times(2)).toString(), "8")
            __check((plus(3)).toString(), "3")
        }
