// vybe-test: kotlin/kotlin_return_expressions/test_return_from_function_with_named_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun join(prefix: String, value: Int): String {
            return "${'$'}prefix${'$'}value"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join(prefix = "x", value = 9)).toString(), "x9")
        }
