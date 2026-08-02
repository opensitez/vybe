// vybe-test: kotlin/numeric_literals/test_decimal_underscore_grouping
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((1_000 + 2_000).toString(), "3000") }
