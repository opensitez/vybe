// vybe-test: kotlin/escaped_identifiers/test_dash_in_identifier
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

fun `value-sum`(a: Int, b: Int): Int = a + b
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((`value-sum`(2, 3)).toString(), "5") }
