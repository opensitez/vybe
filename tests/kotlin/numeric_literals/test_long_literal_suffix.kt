// vybe-test: kotlin/numeric_literals/test_long_literal_suffix
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val v: Long = 12L
__check((v + 3L).toString(), "15") }
