// vybe-test: kotlin/kotlin_expression_body_forms/test_expression_lambda
// origin: languages/kotlin/tests/kotlin/test_kotlin_expression_body_forms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sum = { a: Int, b: Int -> a + b }
            __check((sum(4, 5).toString()).toString(), "9")
        }
