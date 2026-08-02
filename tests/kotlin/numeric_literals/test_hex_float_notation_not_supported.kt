// vybe-test: kotlin/numeric_literals/test_hex_float_notation_not_supported
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((1_000_000).toString(), "1000000") }
