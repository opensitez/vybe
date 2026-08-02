// vybe-test: kotlin/functions/test_function_early_return
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun checkPositive(x: Int) {
            if (x <= 0) return
            __check(("positive").toString(), "positive")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            checkPositive(-5)
            checkPositive(5)
        }
