// vybe-test: kotlin/string_templates/test_template_nullable_value_null
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: String? = null
            __check(("value=${v}").toString(), "value=null")
        }
