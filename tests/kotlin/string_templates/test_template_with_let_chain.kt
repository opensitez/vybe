// vybe-test: kotlin/string_templates/test_template_with_let_chain
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).let { it.size }
            __check(("let=$out").toString(), "let=3")
        }
