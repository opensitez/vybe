// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_escaped_dollar_output
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("literal ${'$'}").toString(), "literal \$")
            __check(("price ${'$'}{10}").toString(), "price 10")
        }
