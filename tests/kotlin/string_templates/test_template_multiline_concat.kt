// vybe-test: kotlin/string_templates/test_template_multiline_concat
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1
            val b = 2
            __check(("a=${a}").toString(), "a=1")
            __check(("sum=${a + b}").toString(), "sum=3")
        }
