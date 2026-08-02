// vybe-test: kotlin/string_templates/test_template_double_quoted_brace
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = 1
            val right = 2
            __check(("${left}+${right}=${left + right}").toString(), "1+2=3")
        }
