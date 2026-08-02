// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_string_template_with_function_call
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun decorate(input: String): String = input.uppercase()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("${'$'}{decorate("kotlin")}").toString(), "KOTLIN")
        }
