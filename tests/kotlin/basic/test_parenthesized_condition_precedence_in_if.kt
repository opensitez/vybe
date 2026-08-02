// vybe-test: kotlin/basic/test_parenthesized_condition_precedence_in_if
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = (true || false) && false
            __check((result).toString(), "false")
            val result2 = true || (false && false)
            __check((result2).toString(), "true")
        }
