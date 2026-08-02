// vybe-test: kotlin/string_templates/test_template_expression_sum
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 2
            val b = 3
            __check(("sum=${a + b}").toString(), "sum=5")
        }
