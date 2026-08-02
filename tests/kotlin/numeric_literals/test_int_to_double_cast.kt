// vybe-test: kotlin/numeric_literals/test_int_to_double_cast
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = 9
__check((x.toDouble() + 0.5).toString(), "9.5") }
