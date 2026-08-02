// vybe-test: kotlin/numeric_literals/test_mixed_float_long
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = 2L + 3.0
__check((x).toString(), "5.0") }
