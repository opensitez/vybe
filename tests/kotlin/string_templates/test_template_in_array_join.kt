// vybe-test: kotlin/string_templates/test_template_in_array_join
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = listOf("x", "y", "z")
            __check(("joined=${v.joinToString(",")}").toString(), "joined=x,y,z")
        }
