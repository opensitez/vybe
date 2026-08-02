// vybe-test: kotlin/string_templates/test_template_when_expression
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 2
            __check(("v=${when(x) { 1 -> "a" 2 -> "b" else -> "c" }}").toString(), "v=b")
        }
