// vybe-test: kotlin/infix/test_infix_expression_chain
// origin: languages/kotlin/tests/kotlin/test_infix.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val keyValue = "a" to "b" to "c"
            val first = keyValue.first
            val second = keyValue.second
            __check((first).toString(), "a")
            __check((second).toString(), "b")
        }
