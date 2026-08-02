// vybe-test: kotlin/string_templates/test_template_boolean_and_math_combo
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = true
            val value = 4
            __check(("ok=${ok} total=${value * 2}").toString(), "ok=true total=8")
        }
