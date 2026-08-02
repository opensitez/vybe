// vybe-test: kotlin/if_expressions/test_if_expression_value_return
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun classify(v: Int): String {
            return if (v > 0) "positive" else if (v < 0) "negative" else "zero"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(3)).toString(), "positive")
            __check((classify(0)).toString(), "zero")
            __check((classify(-1)).toString(), "negative")
        }
