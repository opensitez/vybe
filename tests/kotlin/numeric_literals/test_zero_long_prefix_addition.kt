// vybe-test: kotlin/numeric_literals/test_zero_long_prefix_addition
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x: Long = 0L
__check((x + 1L).toString(), "1") }
