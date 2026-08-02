// vybe-test: kotlin/string_templates/test_template_with_conditional_operator
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = -1
            __check(("isPositive=${if (n > 0) "yes" else "no"}").toString(), "isPositive=no")
        }
