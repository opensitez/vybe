// vybe-test: kotlin/escaped_identifiers/test_function_with_symbols
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun `f+g`(a: Int): Int = a + 10
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`f+g`(1)).toString(), "11") }
