// vybe-test: kotlin/escaped_identifiers/test_space_in_identifier
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun `add one`(x: Int): Int = x + 1
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`add one`(2)).toString(), "3") }
