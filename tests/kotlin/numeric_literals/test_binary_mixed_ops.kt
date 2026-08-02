// vybe-test: kotlin/numeric_literals/test_binary_mixed_ops
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((0b11 * 0x2).toString(), "6") }
