// vybe-test: kotlin/kotlin_operator_overflow/test_char_codepoint_arithmetic
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('A'.code + 1).toString(), "66")
            __check((('A'.code + 1).toChar()).toString(), "B")
        }
