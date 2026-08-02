// vybe-test: kotlin/printing/test_printing_string_template_expression
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = 2
            val right = 5
            __check(("sum=${left + right}").toString(), "sum=7")
            __check(("$left*$right=${left * right}").toString(), "2*5=10")
        }
