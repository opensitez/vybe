// vybe-test: kotlin/kotlin_return_expressions/test_expression_function_syntax
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun square(x: Int): Int = x * x

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((square(4)).toString(), "16")
        }
