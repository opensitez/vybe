// vybe-test: kotlin/numeric_literals/test_long_hex_literal
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val v: Long = 0x10L
__check((v).toString(), "16") }
