// vybe-test: kotlin/numeric_literals/test_binary_literal_small
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((0b1010 + 0b0101).toString(), "15") }
