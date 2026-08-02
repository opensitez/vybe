// vybe-test: kotlin/numeric_literals/test_float_suffix
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val f: Float = 1.25f
__check((f * 2).toString(), "2.5") }
