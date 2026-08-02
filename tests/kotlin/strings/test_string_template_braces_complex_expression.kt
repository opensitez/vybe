// vybe-test: kotlin/strings/test_string_template_braces_complex_expression
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val width = 4
            val height = 5
            __check(("${width}x${height}=${width * height}").toString(), "4x5=20")
            __check(("${'$'}${width + height}").toString(), "\$9")
        }
