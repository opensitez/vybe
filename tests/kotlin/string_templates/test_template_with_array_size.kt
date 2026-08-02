// vybe-test: kotlin/string_templates/test_template_with_array_size
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(1, 2, 3)
            __check(("size=${items.size}").toString(), "size=3")
        }
