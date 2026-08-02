// vybe-test: kotlin/functions/test_function_expression_body_logic
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun maxOf(a: Int, b: Int): Int = if (a > b) a else b

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxOf(10, 20)).toString(), "20")
        }
