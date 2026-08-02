// vybe-test: kotlin/numeric_literals/test_numeric_shift
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = 1 shl 2
__check((x).toString(), "4") }
