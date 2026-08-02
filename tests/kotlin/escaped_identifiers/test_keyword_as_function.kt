// vybe-test: kotlin/escaped_identifiers/test_keyword_as_function
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun `when`(): Int = 5
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`when`()).toString(), "5") }
