// vybe-test: kotlin/strings/test_string_template_expression
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 3
            val b = 4
            __check(("$a + $b = ${a + b}").toString(), "3 + 4 = 7")
        }
