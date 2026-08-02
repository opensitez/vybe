// vybe-test: kotlin/string_templates/test_template_with_indexed_access
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abc"
            __check(("third=${s[2]}").toString(), "third=c")
        }
