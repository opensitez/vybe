// vybe-test: kotlin/numeric_literals/test_literal_range_inclusive
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val r = 1..3
__check((r.start == 1 && r.endInclusive == 3).toString(), "true") }
