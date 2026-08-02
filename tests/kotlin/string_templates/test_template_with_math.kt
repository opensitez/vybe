// vybe-test: kotlin/string_templates/test_template_with_math
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("pi=${kotlin.math.PI}").toString(), "pi=3.141592653589793")
        }
