// vybe-test: kotlin/numeric_literals/test_mixed_hex_and_decimal_addition
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((0x5 + 3).toString(), "8") }
