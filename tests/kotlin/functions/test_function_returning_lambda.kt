// vybe-test: kotlin/functions/test_function_returning_lambda
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun makeMultiplier(mult: Int): (Int) -> Int {
            return { value -> value * mult }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val timesFive = makeMultiplier(5)
            __check((timesFive(3)).toString(), "15")
        }
