// vybe-test: kotlin/string_templates/test_template_nested_calls
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("len=${"abc".length + "de".length}").toString(), "len=5")
        }
