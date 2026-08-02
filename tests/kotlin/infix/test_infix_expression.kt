// vybe-test: kotlin/infix/test_infix_expression
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = "key" to "value"
            val (key, value) = p
            __check((key).toString(), "key")
            __check((value).toString(), "value")
        }
