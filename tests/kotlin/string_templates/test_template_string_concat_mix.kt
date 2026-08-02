// vybe-test: kotlin/string_templates/test_template_string_concat_mix
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = "a"
            val q = "b"
            __check(("$p+$q=${p + q}").toString(), "a+b=ab")
        }
