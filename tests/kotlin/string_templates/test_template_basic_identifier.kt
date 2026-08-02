// vybe-test: kotlin/string_templates/test_template_basic_identifier
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "k"
            __check(("hello $name").toString(), "hello k")
        }
