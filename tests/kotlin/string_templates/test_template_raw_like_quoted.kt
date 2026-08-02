// vybe-test: kotlin/string_templates/test_template_raw_like_quoted
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("path=/tmp/${"a"}").toString(), "path=/tmp/a")
        }
