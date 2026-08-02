// vybe-test: kotlin/string_templates/test_template_block_of_string
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 3
            __check(("value=${if (x % 2 == 0) "even" else "odd"}").toString(), "value=odd")
        }
