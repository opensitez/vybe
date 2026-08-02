// vybe-test: kotlin/numeric_literals/test_long_integer_division
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a: Long = 10L
__check((a / 2L).toString(), "5") }
