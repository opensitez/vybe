// vybe-test: kotlin/functions/test_function_optional_else_return
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun maybePositive(n: Int): String {
            if (n > 0) return "positive"
            return "not positive"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maybePositive(3)).toString(), "positive")
            __check((maybePositive(-1)).toString(), "not positive")
        }
