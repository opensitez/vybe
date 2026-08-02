// vybe-test: kotlin/functions/test_function_expression_body
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun square(x: Int): Int = x * x

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((square(6)).toString(), "36")
        }
