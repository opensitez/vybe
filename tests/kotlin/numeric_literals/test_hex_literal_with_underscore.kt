// vybe-test: kotlin/numeric_literals/test_hex_literal_with_underscore
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((0x1_0 + 0x2).toString(), "18") }
