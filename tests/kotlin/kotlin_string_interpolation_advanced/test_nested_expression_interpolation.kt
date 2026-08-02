// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_nested_expression_interpolation
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            val msg = "value is ${'$'}{if (value > 3) "high" else "low"} and ${'$'}value"
            __check((msg).toString(), "value is high and 5")
        }
