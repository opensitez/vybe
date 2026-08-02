// vybe-test: kotlin/string_templates/test_template_joiner_with_nulls
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: List<String?> = listOf("a", null, "b")
            __check(("join=${values.joinToString(",")}").toString(), "join=a,null,b")
        }
