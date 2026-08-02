// vybe-test: kotlin/string_templates/test_template_unicode_literal
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("heart=\u2665").toString(), "heart=♥")
        }
