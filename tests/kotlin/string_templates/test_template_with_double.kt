// vybe-test: kotlin/string_templates/test_template_with_double
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = 1.5
            __check(("double=$d").toString(), "double=1.5")
        }
