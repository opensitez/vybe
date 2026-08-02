// vybe-test: kotlin/kotlin_return_expressions/test_early_return_in_expression_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun runWithGuard(v: Int): String {
            return run {
                if (v == 0) return@run "zero"
                "value"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((runWithGuard(0)).toString(), "zero")
            __check((runWithGuard(2)).toString(), "value")
        }
