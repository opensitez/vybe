// vybe-test: kotlin/numeric_literals/test_unsigned_not_supported_fallback
// origin: languages/kotlin/tests/kotlin/test_numeric_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a: Long = 10L
__check((a).toString(), "10") }
