// vybe-test: kotlin/numeric_literals/test_string_length_from_length_literal
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val n = "123456".length
__check((n).toString(), "6") }
