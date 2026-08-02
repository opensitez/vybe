// vybe-test: kotlin/literals/test_string_template_with_expression
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 2
            val b = 4
            __check(("$a + $b = ${a + b}").toString(), "2 + 4 = 6")
        }
