// vybe-test: kotlin/invoke_operator/test_invoke_nested_object_expression
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val factory = object {
                operator fun invoke(a: Int, b: Int): Int = a + b
            }
            __check((factory(4, 5)).toString(), "9")
        }
