// vybe-test: kotlin/numeric_literals/test_double_suffix
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val d: Double = 2.0
__check((d * 3).toString(), "6.0") }
