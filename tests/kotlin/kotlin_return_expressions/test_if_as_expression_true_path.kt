// vybe-test: kotlin/kotlin_return_expressions/test_if_as_expression_true_path
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun classify(v: Int): String {
            return if (v > 0) "positive" else "non-positive"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(1)).toString(), "positive")
            __check((classify(0)).toString(), "non-positive")
        }
